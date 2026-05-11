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

        private QueryTabPage ActiveTab => (QueryTabPage)tabQueries.SelectedTab;

        public MainForm()
        {
            InitializeComponent();
            UpdateConnectionStatus(false);
            AddNewTab();

            Load += (s, e) =>
            {
                splitOuter.Panel1MinSize = 400;
                splitOuter.Panel2MinSize = 180;
                splitOuter.SplitterDistance = (int)(splitOuter.Width * 0.75);
            };

            KeyPreview = true;
            KeyDown += MainForm_KeyDown;
        }

        // ── Tab management ──────────────────────────────────────────────────────

        private void AddNewTab()
        {
            var tab = new QueryTabPage();
            tab.Editor.KeyDown += rtbQuery_KeyDown;
            tabQueries.TabPages.Add(tab);
            tabQueries.SelectedTab = tab;
            tab.Editor.Focus();
        }

        private void CloseActiveTab()
        {
            if (tabQueries.TabPages.Count <= 1) return;
            var current = tabQueries.SelectedTab;
            int idx = tabQueries.SelectedIndex;
            tabQueries.TabPages.Remove(current);
            tabQueries.SelectedIndex = Math.Min(idx, tabQueries.TabPages.Count - 1);
        }

        private void btnNewTab_Click(object sender, EventArgs e) => AddNewTab();

        // Right-click tab → context menu to close
        private void tabQueries_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right && e.Button != MouseButtons.Middle) return;

            for (int i = 0; i < tabQueries.TabPages.Count; i++)
            {
                if (!tabQueries.GetTabRect(i).Contains(e.Location)) continue;
                tabQueries.SelectedIndex = i;

                if (e.Button == MouseButtons.Middle)
                {
                    CloseActiveTab();
                    return;
                }

                var menu = new ContextMenuStrip();
                menu.Items.Add("Close Tab", null, (s, _) => CloseActiveTab());
                if (tabQueries.TabPages.Count > 1)
                    menu.Items.Add("Close Other Tabs", null, (s, _) =>
                    {
                        var keep = tabQueries.SelectedTab;
                        for (int j = tabQueries.TabPages.Count - 1; j >= 0; j--)
                            if (tabQueries.TabPages[j] != keep)
                                tabQueries.TabPages.RemoveAt(j);
                    });
                menu.Items.Add("-");
                menu.Items.Add("Rename Tab", null, (s, _) =>
                {
                    string? name = SimplePrompt("Rename Tab", "Tab name:", ActiveTab.Text);
                    if (name != null) ActiveTab.Text = name;
                });
                menu.Show(tabQueries, e.Location);
                return;
            }
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.W)
            {
                e.SuppressKeyPress = true;
                CloseActiveTab();
            }
            if (e.Control && e.KeyCode == Keys.T)
            {
                e.SuppressKeyPress = true;
                AddNewTab();
            }
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
                var b = new SqlConnectionStringBuilder(_connectionString);
                lblConnInfo.Text = $"{b.DataSource}  |  {b.InitialCatalog}";
                UpdateConnectionStatus(true);
                ActiveTab.AppendMessage($"Connected to {b.DataSource} / {b.InitialCatalog}", Color.Green);
            }
            catch (Exception ex)
            {
                _isConnected = false;
                UpdateConnectionStatus(false);
                ActiveTab.AppendMessage($"Connection failed: {ex.Message}", Color.Red);
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Execute ─────────────────────────────────────────────────────────────

        private async void btnExecute_Click(object sender, EventArgs e)
        {
            if (!_isConnected)
            {
                MessageBox.Show("Please connect to a database first.", "Not Connected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string sql = ActiveTab.GetActiveQuery();
            if (string.IsNullOrWhiteSpace(sql)) return;

            var tab = ActiveTab;
            SetExecutingState(true);
            tab.ShowResultsTab();

            var queryTask = Task.Run(() => ExecuteQuery(sql));
            var columnsTask = FetchColumnsFromQueryAsync(sql);

            try
            {
                var sw = Stopwatch.StartNew();
                await Task.WhenAll(queryTask, columnsTask);
                sw.Stop();

                var result = await queryTask;
                var dt = result.Tables.Count > 0 ? result.Tables[0] : null;
                tab.SetResult(dt, sw.ElapsedMilliseconds);
                tab.AppendMessage($"Query executed in {sw.ElapsedMilliseconds} ms.", Color.DodgerBlue);

                PopulateColumnBrowserFromQuery(await columnsTask, sql);
            }
            catch (Exception ex)
            {
                tab.ShowMessagesTab();
                tab.AppendMessage($"Error: {ex.Message}", Color.Red);
                try { PopulateColumnBrowserFromQuery(await columnsTask, sql); } catch { }
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
            string sql = ActiveTab.GetActiveQuery();
            if (string.IsNullOrWhiteSpace(sql)) return;

            if (MessageBox.Show("Execute non-query (INSERT/UPDATE/DELETE/DDL)?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            var tab = ActiveTab;
            SetExecutingState(true);
            tab.ShowMessagesTab();

            var columnsTask = FetchColumnsFromQueryAsync(sql);
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
                tab.AppendMessage($"{rows} row(s) affected  ({sw.ElapsedMilliseconds} ms).", Color.DodgerBlue);
                tab.LblRowCount.Text = $"{rows} row(s) affected  |  {sw.ElapsedMilliseconds} ms";
                PopulateColumnBrowserFromQuery(await columnsTask, sql);
            }
            catch (Exception ex)
            {
                tab.AppendMessage($"Error: {ex.Message}", Color.Red);
                try { PopulateColumnBrowserFromQuery(await columnsTask, sql); } catch { }
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        // ── Column browser ──────────────────────────────────────────────────────

        // Called after Execute: parse tables from SQL, fetch columns
        private async Task<List<(string Table, string Column, string DataType, string Nullable)>>
            FetchColumnsFromQueryAsync(string sql)
        {
            var tables = ParseTableNames(sql);
            if (tables.Count == 0 || !_isConnected)
                return new();
            return await Task.Run(() => FetchColumns(tables));
        }

        // Called from search box Enter key: fetch all columns of the given table name
        private async void txtTableSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;

            string tableName = txtTableSearch.Text.Trim();
            if (string.IsNullOrEmpty(tableName)) { ClearColumnBrowser(); return; }
            if (!_isConnected)
            {
                lblColumnBrowserStatus.Text = "Not connected";
                return;
            }

            lblColumnBrowserStatus.Text = $"Searching \"{tableName}\"...";
            treeColumns.Nodes.Clear();

            try
            {
                // Wildcard: if user typed "%" treat as LIKE, otherwise exact match
                var columns = await Task.Run(() => FetchColumnsBySearch(tableName));
                PopulateColumnBrowserDirect(columns);
            }
            catch (Exception ex)
            {
                lblColumnBrowserStatus.Text = $"Error: {ex.Message}";
            }
        }

        private List<(string Table, string Column, string DataType, string Nullable)>
            FetchColumnsBySearch(string tableNamePattern)
        {
            bool isWild = tableNamePattern.Contains('%') || tableNamePattern.Contains('_');
            string op = isWild ? "LIKE" : "=";

            string query = $@"
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, IS_NULLABLE
FROM   INFORMATION_SCHEMA.COLUMNS
WHERE  TABLE_NAME {op} @pattern
ORDER  BY TABLE_NAME, ORDINAL_POSITION";

            var result = new List<(string, string, string, string)>();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(query, conn) { CommandTimeout = 30 };
            cmd.Parameters.AddWithValue("@pattern", tableNamePattern);
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
                result.Add((reader.GetString(0), reader.GetString(1),
                            reader.GetString(2), reader.GetString(3)));
            return result;
        }

        private List<(string Table, string Column, string DataType, string Nullable)>
            FetchColumns(List<string> tableNames)
        {
            var paramNames = new List<string>();
            for (int i = 0; i < tableNames.Count; i++) paramNames.Add($"@t{i}");
            string inClause = string.Join(",", paramNames);

            string query = $@"
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE, IS_NULLABLE
FROM   INFORMATION_SCHEMA.COLUMNS
WHERE  TABLE_NAME IN ({inClause})
ORDER  BY TABLE_NAME, ORDINAL_POSITION";

            var result = new List<(string, string, string, string)>();
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

        private static List<string> ParseTableNames(string sql)
        {
            var tables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var pattern = new Regex(
                @"(?:FROM|JOIN)\s+(\[?[\w]+\]?\.)?(\[?([\w]+)\]?)",
                RegexOptions.IgnoreCase | RegexOptions.Multiline);
            foreach (Match m in pattern.Matches(sql))
            {
                string name = m.Groups[3].Value.Trim('[', ']');
                if (!string.Equals(name, "SELECT", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(name, "WITH", StringComparison.OrdinalIgnoreCase))
                    tables.Add(name);
            }
            return new List<string>(tables);
        }

        private void PopulateColumnBrowserFromQuery(
            List<(string Table, string Column, string DataType, string Nullable)> columns,
            string sql)
        {
            if (InvokeRequired) { Invoke(() => PopulateColumnBrowserFromQuery(columns, sql)); return; }

            treeColumns.BeginUpdate();
            treeColumns.Nodes.Clear();

            if (columns.Count == 0)
            {
                var tables = ParseTableNames(sql);
                lblColumnBrowserStatus.Text = tables.Count == 0
                    ? "No tables found in query"
                    : $"No schema info for: {string.Join(", ", tables)}";
                treeColumns.EndUpdate();
                return;
            }

            BuildTreeNodes(columns);
            treeColumns.ExpandAll();
            treeColumns.EndUpdate();

            var unique = new HashSet<string>();
            foreach (var (t, _, _, _) in columns) unique.Add(t);
            lblColumnBrowserStatus.Text = $"{unique.Count} table(s)  |  {columns.Count} column(s)";
        }

        private void PopulateColumnBrowserDirect(
            List<(string Table, string Column, string DataType, string Nullable)> columns)
        {
            if (InvokeRequired) { Invoke(() => PopulateColumnBrowserDirect(columns)); return; }

            treeColumns.BeginUpdate();
            treeColumns.Nodes.Clear();

            if (columns.Count == 0)
            {
                lblColumnBrowserStatus.Text = "Table not found";
                treeColumns.EndUpdate();
                return;
            }

            BuildTreeNodes(columns);
            treeColumns.ExpandAll();
            treeColumns.EndUpdate();

            var unique = new HashSet<string>();
            foreach (var (t, _, _, _) in columns) unique.Add(t);
            lblColumnBrowserStatus.Text = $"{unique.Count} table(s)  |  {columns.Count} column(s)";
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
                    Tag = column
                };
                tableNode!.Nodes.Add(colNode);
            }
        }

        private void ClearColumnBrowser()
        {
            treeColumns.Nodes.Clear();
            lblColumnBrowserStatus.Text = "Execute a query to see columns";
        }

        // Double-click column → insert name at cursor
        private void treeColumns_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string colName && tabQueries.SelectedTab is QueryTabPage tab)
            {
                int sel = tab.Editor.SelectionStart;
                tab.Editor.Text = tab.Editor.Text.Insert(sel, colName);
                tab.Editor.SelectionStart = sel + colName.Length;
                tab.Editor.Focus();
            }
        }

        // ── Keyboard shortcuts ──────────────────────────────────────────────────

        private void rtbQuery_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnExecute_Click(sender, EventArgs.Empty);
            }
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private void UpdateConnectionStatus(bool connected)
        {
            picStatus.BackColor = connected ? Color.LimeGreen : Color.Gray;
            btnDisconnect.Enabled = connected;
            btnExecute.Enabled = connected;
            btnExecuteNonQuery.Enabled = connected;
        }

        private void SetExecutingState(bool executing)
        {
            btnExecute.Enabled = !executing;
            btnExecuteNonQuery.Enabled = !executing;
            progressBar.Visible = executing;
            lblStatus.Text = executing ? "Executing..." : "Ready";
        }

        private static string? SimplePrompt(string title, string label, string defaultValue)
        {
            var frm = new Form
            {
                Text = title, Size = new Size(340, 130),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };
            var lbl = new Label { Text = label, Location = new Point(10, 12), AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9.5f) };
            var txt = new TextBox { Text = defaultValue, Location = new Point(10, 32), Width = 304,
                BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9.5f) };
            var ok = new Button { Text = "OK", Location = new Point(148, 64), Width = 75, Height = 28,
                DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White };
            ok.FlatAppearance.BorderSize = 0;
            var cancel = new Button { Text = "Cancel", Location = new Point(233, 64), Width = 75, Height = 28,
                DialogResult = DialogResult.Cancel, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72), ForeColor = Color.White };
            cancel.FlatAppearance.BorderSize = 0;
            frm.AcceptButton = ok; frm.CancelButton = cancel;
            frm.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
            txt.SelectAll();
            return frm.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : null;
        }
    }
}
