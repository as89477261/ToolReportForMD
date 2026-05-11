using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace ReportLineOAForDebtAndBranch
{
    public partial class MainForm : Form
    {
        private string _connectionString = string.Empty;
        private bool _isConnected = false;

        // Stores the full column data for filter support
        private readonly List<(string Table, string Column, string DataType, string Nullable)> _columnCache = new();

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
            ClearColumnBrowser();
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

            // Run query and column browser fetch in parallel
            var queryTask = Task.Run(() => ExecuteQuery(sql));
            var columnsTask = FetchColumnsForQueryAsync(sql);

            try
            {
                var sw = Stopwatch.StartNew();
                await Task.WhenAll(queryTask, columnsTask);
                sw.Stop();

                var result = await queryTask;
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

                PopulateColumnBrowser(await columnsTask, sql);
            }
            catch (Exception ex)
            {
                tabResults.SelectedTab = tabPageMessages;
                AppendMessage($"Error: {ex.Message}", Color.Red);

                // Still try to show columns even if query failed
                try { PopulateColumnBrowser(await columnsTask, sql); } catch { }
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

            var columnsTask = FetchColumnsForQueryAsync(sql);
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

                PopulateColumnBrowser(await columnsTask, sql);
            }
            catch (Exception ex)
            {
                AppendMessage($"Error: {ex.Message}", Color.Red);
                try { PopulateColumnBrowser(await columnsTask, sql); } catch { }
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        // ── Column Browser ──────────────────────────────────────────────────────

        private static List<string> ParseTableNames(string sql)
        {
            var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Match: FROM TableName / JOIN TableName / FROM [schema].[Table] / FROM schema.Table
            var pattern = new Regex(
                @"(?:FROM|JOIN)\s+(\[?[\w]+\]?\.)?(\[?([\w]+)\]?)",
                RegexOptions.IgnoreCase | RegexOptions.Multiline);

            foreach (Match m in pattern.Matches(sql))
            {
                string name = m.Groups[3].Value.Trim('[', ']');
                // Skip SQL keywords that can appear after FROM/JOIN
                if (!string.Equals(name, "SELECT", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(name, "WITH", StringComparison.OrdinalIgnoreCase))
                    tables.Add(name);
            }

            return new List<string>(tables);
        }

        private async Task<List<(string Table, string Column, string DataType, string Nullable)>>
            FetchColumnsForQueryAsync(string sql)
        {
            var tableNames = ParseTableNames(sql);
            if (tableNames.Count == 0 || !_isConnected)
                return new List<(string, string, string, string)>();

            return await Task.Run(() => FetchColumns(tableNames));
        }

        private List<(string Table, string Column, string DataType, string Nullable)>
            FetchColumns(List<string> tableNames)
        {
            var result = new List<(string, string, string, string)>();

            // Build parameterised IN list
            var paramNames = new List<string>();
            for (int i = 0; i < tableNames.Count; i++)
                paramNames.Add($"@t{i}");

            string inClause = string.Join(",", paramNames);
            string query = $@"
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, IS_NULLABLE
FROM   INFORMATION_SCHEMA.COLUMNS
WHERE  TABLE_NAME IN ({inClause})
ORDER  BY TABLE_NAME, ORDINAL_POSITION";

            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(query, conn) { CommandTimeout = 30 };
            for (int i = 0; i < tableNames.Count; i++)
                cmd.Parameters.AddWithValue(paramNames[i], tableNames[i]);

            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add((reader.GetString(0), reader.GetString(1),
                            reader.GetString(2), reader.GetString(3)));

            return result;
        }

        private void PopulateColumnBrowser(
            List<(string Table, string Column, string DataType, string Nullable)> columns,
            string sql)
        {
            if (InvokeRequired)
            {
                Invoke(() => PopulateColumnBrowser(columns, sql));
                return;
            }

            _columnCache.Clear();
            _columnCache.AddRange(columns);

            treeColumns.BeginUpdate();
            treeColumns.Nodes.Clear();
            txtTableFilter.Clear();

            if (columns.Count == 0)
            {
                var tables = ParseTableNames(sql);
                lblColumnBrowserStatus.Text = tables.Count == 0
                    ? "No tables found in query"
                    : $"No columns found for: {string.Join(", ", tables)}";
                treeColumns.EndUpdate();
                return;
            }

            BuildTreeNodes(columns);
            treeColumns.ExpandAll();
            treeColumns.EndUpdate();

            var uniqueTables = new HashSet<string>();
            foreach (var (t, _, _, _) in columns) uniqueTables.Add(t);
            lblColumnBrowserStatus.Text = $"{uniqueTables.Count} table(s)  |  {columns.Count} column(s)";
        }

        private void BuildTreeNodes(
            IEnumerable<(string Table, string Column, string DataType, string Nullable)> columns)
        {
            string? currentTable = null;
            TreeNode? tableNode = null;

            foreach (var (table, column, dataType, nullable) in columns)
            {
                if (table != currentTable)
                {
                    tableNode = new TreeNode($"[{table}]")
                    {
                        ForeColor = Color.FromArgb(86, 156, 214),
                        NodeFont = new Font("Consolas", 9f, FontStyle.Bold)
                    };
                    treeColumns.Nodes.Add(tableNode);
                    currentTable = table;
                }

                string nullMark = nullable == "YES" ? "?" : "";
                var colNode = new TreeNode($"{column}  ({dataType}{nullMark})")
                {
                    ForeColor = Color.FromArgb(212, 212, 212),
                    Tag = column   // store raw name for double-click insert
                };
                tableNode!.Nodes.Add(colNode);
            }
        }

        private void ClearColumnBrowser()
        {
            _columnCache.Clear();
            treeColumns.Nodes.Clear();
            lblColumnBrowserStatus.Text = "Execute a query to see columns";
        }

        // Filter tree as user types
        private void txtTableFilter_TextChanged(object sender, EventArgs e)
        {
            string filter = txtTableFilter.Text.Trim();

            treeColumns.BeginUpdate();
            treeColumns.Nodes.Clear();

            var filtered = string.IsNullOrEmpty(filter)
                ? _columnCache
                : _columnCache.FindAll(c =>
                    c.Column.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                    c.Table.Contains(filter, StringComparison.OrdinalIgnoreCase));

            if (filtered.Count > 0)
                BuildTreeNodes(filtered);

            treeColumns.ExpandAll();
            treeColumns.EndUpdate();
        }

        // Double-click a column node → insert column name at cursor in query editor
        private void treeColumns_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string colName)
            {
                int sel = rtbQuery.SelectionStart;
                rtbQuery.Text = rtbQuery.Text.Insert(sel, colName);
                rtbQuery.SelectionStart = sel + colName.Length;
                rtbQuery.Focus();
            }
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private string GetActiveQuery()
        {
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
            ClearColumnBrowser();
        }

        private void btnClearQuery_Click(object sender, EventArgs e) => rtbQuery.Clear();

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
