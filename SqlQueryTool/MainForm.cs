using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace SqlQueryTool
{
    public partial class MainForm : Form
    {
        private string _connectionString = string.Empty;
        private bool _isConnected = false;

        public MainForm()
        {
            InitializeComponent();
            UpdateConnectionStatus(false);
        }

        // ── Connection ──────────────────────────────────────────────────────────

        private void btnConnect_Click(object sender, EventArgs e)
        {
            using var dlg = new ConnectionDialog(_connectionString);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                _connectionString = dlg.ConnectionString;
                TestAndConnect();
            }
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            _connectionString = string.Empty;
            _isConnected = false;
            UpdateConnectionStatus(false);
            lblConnInfo.Text = "Not connected";
        }

        private void TestAndConnect()
        {
            try
            {
                using var conn = new SqlConnection(_connectionString);
                conn.Open();
                _isConnected = true;

                var builder = new SqlConnectionStringBuilder(_connectionString);
                lblConnInfo.Text = $"{builder.DataSource}  |  {builder.InitialCatalog}";
                UpdateConnectionStatus(true);
                AppendMessage($"Connected to {builder.DataSource} / {builder.InitialCatalog}", Color.Green);
            }
            catch (Exception ex)
            {
                _isConnected = false;
                UpdateConnectionStatus(false);
                AppendMessage($"Connection failed: {ex.Message}", Color.Red);
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Execute Query ───────────────────────────────────────────────────────

        private async void btnExecute_Click(object sender, EventArgs e)
        {
            if (!_isConnected)
            {
                MessageBox.Show("Please connect to a database first.", "Not Connected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sql = GetActiveQuery();
            if (string.IsNullOrWhiteSpace(sql)) return;

            SetExecutingState(true);
            tabResults.SelectedTab = tabPageResults;

            try
            {
                var sw = Stopwatch.StartNew();
                var result = await Task.Run(() => ExecuteQuery(sql));
                sw.Stop();

                if (result.Tables.Count > 0 && result.Tables[0].Rows.Count > 0)
                {
                    grid.DataSource = result.Tables[0];
                    lblRowCount.Text = $"{result.Tables[0].Rows.Count} row(s)";
                }
                else
                {
                    grid.DataSource = null;
                    lblRowCount.Text = "0 row(s)";
                }

                AppendMessage($"Query executed in {sw.ElapsedMilliseconds} ms.", Color.DodgerBlue);
                lblExecTime.Text = $"{sw.ElapsedMilliseconds} ms";
            }
            catch (Exception ex)
            {
                tabResults.SelectedTab = tabPageMessages;
                AppendMessage($"Error: {ex.Message}", Color.Red);
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        private DataSet ExecuteQuery(string sql)
        {
            var ds = new DataSet();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(sql, conn) { CommandTimeout = 60 };
            using var adapter = new SqlDataAdapter(cmd);
            adapter.Fill(ds);
            return ds;
        }

        private async void btnExecuteNonQuery_Click(object sender, EventArgs e)
        {
            if (!_isConnected) return;

            string sql = GetActiveQuery();
            if (string.IsNullOrWhiteSpace(sql)) return;

            if (MessageBox.Show("Execute non-query (INSERT/UPDATE/DELETE/DDL)?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            SetExecutingState(true);
            tabResults.SelectedTab = tabPageMessages;

            try
            {
                var sw = Stopwatch.StartNew();
                int rows = await Task.Run(() =>
                {
                    using var conn = new SqlConnection(_connectionString);
                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn) { CommandTimeout = 60 };
                    return cmd.ExecuteNonQuery();
                });
                sw.Stop();

                AppendMessage($"{rows} row(s) affected  ({sw.ElapsedMilliseconds} ms).", Color.DodgerBlue);
                lblRowCount.Text = $"{rows} row(s) affected";
                lblExecTime.Text = $"{sw.ElapsedMilliseconds} ms";
            }
            catch (Exception ex)
            {
                AppendMessage($"Error: {ex.Message}", Color.Red);
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private string GetActiveQuery()
        {
            // If user selected text, run only the selection
            string selected = rtbQuery.SelectedText.Trim();
            return string.IsNullOrEmpty(selected) ? rtbQuery.Text.Trim() : selected;
        }

        private void UpdateConnectionStatus(bool connected)
        {
            picStatus.BackColor = connected ? Color.LimeGreen : Color.Gray;
            btnDisconnect.Enabled = connected;
            btnExecute.Enabled = connected;
            btnExecuteNonQuery.Enabled = connected;
            btnClearResults.Enabled = true;
        }

        private void SetExecutingState(bool executing)
        {
            btnExecute.Enabled = !executing;
            btnExecuteNonQuery.Enabled = !executing;
            progressBar.Visible = executing;
            lblStatus.Text = executing ? "Executing..." : "Ready";
        }

        private void AppendMessage(string text, Color color)
        {
            if (InvokeRequired) { Invoke(() => AppendMessage(text, color)); return; }

            int start = rtbMessages.TextLength;
            string line = $"[{DateTime.Now:HH:mm:ss}]  {text}{Environment.NewLine}";
            rtbMessages.AppendText(line);
            rtbMessages.Select(start, line.Length);
            rtbMessages.SelectionColor = color;
            rtbMessages.SelectionLength = 0;
            rtbMessages.ScrollToCaret();
        }

        private void btnClearResults_Click(object sender, EventArgs e)
        {
            grid.DataSource = null;
            rtbMessages.Clear();
            lblRowCount.Text = string.Empty;
            lblExecTime.Text = string.Empty;
        }

        private void btnClearQuery_Click(object sender, EventArgs e) => rtbQuery.Clear();

        // Ctrl+Enter shortcut in query box
        private void rtbQuery_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnExecute_Click(sender, EventArgs.Empty);
            }
        }

        // Export result to CSV
        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            if (grid.DataSource is not DataTable dt || dt.Rows.Count == 0) return;

            using var dlg = new SaveFileDialog
            {
                Filter = "CSV files|*.csv",
                FileName = "result.csv"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                var lines = new List<string>();
                var headers = new List<string>();
                foreach (DataColumn col in dt.Columns) headers.Add(CsvEscape(col.ColumnName));
                lines.Add(string.Join(",", headers));

                foreach (DataRow row in dt.Rows)
                {
                    var cells = new List<string>();
                    foreach (var item in row.ItemArray)
                        cells.Add(CsvEscape(item?.ToString() ?? string.Empty));
                    lines.Add(string.Join(",", cells));
                }

                System.IO.File.WriteAllLines(dlg.FileName, lines, System.Text.Encoding.UTF8);
                AppendMessage($"Exported {dt.Rows.Count} row(s) to {dlg.FileName}", Color.Green);
            }
            catch (Exception ex)
            {
                AppendMessage($"Export error: {ex.Message}", Color.Red);
            }
        }

        private static string CsvEscape(string s)
        {
            if (s.Contains(',') || s.Contains('"') || s.Contains('\n'))
                return $"\"{s.Replace("\"", "\"\"")}\"";
            return s;
        }
    }
}
