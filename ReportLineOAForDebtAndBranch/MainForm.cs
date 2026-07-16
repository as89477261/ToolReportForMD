using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;

namespace ReportLineOAForDebtAndBranch
{
    public partial class MainForm : Form
    {
        private string _connectionString = string.Empty;
        private bool _isConnected = false;

        // In-memory history list (pinned first, then unpinned newest-first)
        private List<HistoryEntry> _history = new();
        private string _historyFilter = string.Empty;

        private QueryTabPage ActiveTab => (QueryTabPage)tabQueries.SelectedTab;

        public MainForm()
        {
            InitializeComponent();
            UpdateConnectionStatus(false);
            RestoreSession();
            LoadHistory();

            Load += (s, e) =>
            {
                splitOuter.Panel1MinSize = 400;
                splitOuter.Panel2MinSize = 200;
                splitOuter.SplitterDistance = Math.Max(400, splitOuter.Width - 260);
            };

            FormClosing += (s, e) => SaveSession();
            KeyPreview   = true;
            KeyDown      += MainForm_KeyDown;
        }

        // ── Session save / restore ──────────────────────────────────────────────

        private void SaveSession()
        {
            var sessions = new List<TabSession>();
            foreach (TabPage tp in tabQueries.TabPages)
                if (tp is QueryTabPage qt)
                    sessions.Add(new TabSession(qt.Text, qt.Editor.Text));
            SessionStore.Save(sessions);
        }

        private void RestoreSession()
        {
            var sessions = SessionStore.Load();
            if (sessions.Count == 0) { AddNewTab(); return; }
            foreach (var s in sessions)
            {
                var tab = CreateTab(s.Name);
                tab.Editor.Text = s.QueryText;
            }
            tabQueries.SelectedIndex = 0;
        }

        // ── Tab management ──────────────────────────────────────────────────────

        private QueryTabPage CreateTab(string name)
        {
            var tab = new QueryTabPage(name);
            tab.Editor.KeyDown += rtbQuery_KeyDown;
            tabQueries.TabPages.Add(tab);
            return tab;
        }

        private void AddNewTab()
        {
            var tab = CreateTab($"Query {tabQueries.TabPages.Count + 1}");
            tabQueries.SelectedTab = tab;
            tab.Editor.Focus();
        }

        private void CloseActiveTab()
        {
            if (tabQueries.TabPages.Count <= 1) return;
            int idx = tabQueries.SelectedIndex;
            tabQueries.TabPages.Remove(tabQueries.SelectedTab);
            tabQueries.SelectedIndex = Math.Min(idx, tabQueries.TabPages.Count - 1);
        }

        private void btnNewTab_Click(object sender, EventArgs e) => AddNewTab();

        private void tabQueries_MouseClick(object sender, MouseEventArgs e)
        {
            for (int i = 0; i < tabQueries.TabPages.Count; i++)
            {
                if (!tabQueries.GetTabRect(i).Contains(e.Location)) continue;
                tabQueries.SelectedIndex = i;

                if (e.Button == MouseButtons.Middle)
                {
                    CloseActiveTab();
                    return;
                }
                if (e.Button == MouseButtons.Right)
                {
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
                }
                return;
            }
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.W) { e.SuppressKeyPress = true; CloseActiveTab(); }
            if (e.Control && e.KeyCode == Keys.T) { e.SuppressKeyPress = true; AddNewTab(); }
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

            try
            {
                var sw = Stopwatch.StartNew();
                var result = await Task.Run(() => ExecuteQuery(sql));
                sw.Stop();

                var dt = result.Tables.Count > 0 ? result.Tables[0] : null;
                tab.SetResult(dt, sw.ElapsedMilliseconds);
                tab.AppendMessage($"Query executed in {sw.ElapsedMilliseconds} ms.", Color.DodgerBlue);

                AddHistoryEntry(sql);
            }
            catch (Exception ex)
            {
                tab.ShowMessagesTab();
                tab.AppendMessage($"Error: {ex.Message}", Color.Red);
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        private int TimeoutSeconds => (int)numTimeout.Value;

        private DataSet ExecuteQuery(string sql)
        {
            var ds = new DataSet();
            using var conn = new SqlConnection(_connectionString);
            conn.Open();
            using var cmd = new SqlCommand(sql, conn) { CommandTimeout = TimeoutSeconds };
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

            try
            {
                var sw = Stopwatch.StartNew();
                int rows = await Task.Run(() =>
                {
                    using var conn = new SqlConnection(_connectionString);
                    conn.Open();
                    using var cmd = new SqlCommand(sql, conn) { CommandTimeout = TimeoutSeconds };
                    return cmd.ExecuteNonQuery();
                });
                sw.Stop();
                tab.AppendMessage($"{rows} row(s) affected  ({sw.ElapsedMilliseconds} ms).", Color.DodgerBlue);
                tab.LblRowCount.Text = $"{rows} row(s) affected  |  {sw.ElapsedMilliseconds} ms";
                AddHistoryEntry(sql);
            }
            catch (Exception ex)
            {
                tab.AppendMessage($"Error: {ex.Message}", Color.Red);
            }
            finally
            {
                SetExecutingState(false);
            }
        }

        // ── History ─────────────────────────────────────────────────────────────

        private void LoadHistory()
        {
            _history = HistoryStore.Load();
            RefreshHistoryList();
        }

        private void AddHistoryEntry(string sql)
        {
            var entry = HistoryStore.AddEntry(sql);
            // Remove duplicate (same id) if exists, then insert at correct position
            _history.RemoveAll(e => e.Id == entry.Id);
            // Unpinned go at front of unpinned section
            int insertAt = _history.FindIndex(e => !e.IsPinned);
            if (insertAt < 0) _history.Add(entry);
            else              _history.Insert(insertAt, entry);

            RefreshHistoryList();
        }

        private void RefreshHistoryList()
        {
            if (InvokeRequired) { Invoke(RefreshHistoryList); return; }

            lvHistory.BeginUpdate();
            lvHistory.Items.Clear();

            string filter = _historyFilter.ToLower();

            // Pinned first, then unpinned
            var pinned   = _history.FindAll(h => h.IsPinned);
            var unpinned = _history.FindAll(h => !h.IsPinned);
            var ordered  = new List<HistoryEntry>(pinned);
            ordered.AddRange(unpinned);

            foreach (var entry in ordered)
            {
                bool hasLabel = !string.IsNullOrEmpty(entry.Label);

                // Filter: match label OR sql
                if (!string.IsNullOrEmpty(filter) &&
                    !entry.Sql.ToLower().Contains(filter) &&
                    !(entry.Label?.ToLower().Contains(filter) ?? false)) continue;

                string pin     = entry.IsPinned ? "📌 " : "    ";
                string display = hasLabel
                    ? $"{pin}🏷 {entry.Label}"
                    : pin + entry.Sql.Trim().Replace("\r\n", " ").Replace("\n", " ") is var p
                        ? (p.Length > 80 ? p[..80] + "…" : p)
                        : "";

                // ToolTip shows full SQL always
                string tooltip = hasLabel
                    ? $"{entry.Label}\n\n{entry.Sql}"
                    : entry.Sql;

                var color = hasLabel
                    ? Color.FromArgb(140, 210, 255)          // light blue = has label
                    : entry.IsPinned
                        ? Color.FromArgb(255, 200, 50)        // gold = pinned
                        : Color.FromArgb(212, 212, 212);      // normal

                var item = new ListViewItem(display)
                {
                    Tag = entry, ToolTipText = tooltip, ForeColor = color
                };
                item.SubItems.Add(HistoryStore.RelativeTime(entry.ExecutedAt));
                lvHistory.Items.Add(item);
            }

            lvHistory.EndUpdate();
        }

        // Single-click → load query into active editor
        private void lvHistory_MouseClick(object sender, MouseEventArgs e)
        {
            var hit = lvHistory.HitTest(e.Location);
            if (hit.Item == null) return;

            if (e.Button == MouseButtons.Left && hit.Item.Tag is HistoryEntry entry)
            {
                ActiveTab.Editor.Text = entry.Sql;
                ActiveTab.Editor.SelectionStart = entry.Sql.Length;
                ActiveTab.Editor.Focus();
            }

            if (e.Button == MouseButtons.Right && hit.Item.Tag is HistoryEntry rEntry)
            {
                hit.Item.Selected = true;
                ShowHistoryContextMenu(rEntry, e.Location);
            }
        }

        private void lvHistory_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var hit = lvHistory.HitTest(e.Location);
            if (hit.Item?.Tag is HistoryEntry entry)
            {
                // Double-click → open in a new tab
                var tab = CreateTab($"Query {tabQueries.TabPages.Count + 1}");
                tab.Editor.Text = entry.Sql;
                tab.Editor.SelectionStart = entry.Sql.Length;
                tabQueries.SelectedTab = tab;
                tab.Editor.Focus();
            }
        }

        private void ShowHistoryContextMenu(HistoryEntry entry, Point location)
        {
            var menu = new ContextMenuStrip();

            menu.Items.Add(entry.IsPinned ? "Unpin" : "📌 Pin", null, (s, _) =>
            {
                HistoryStore.SetPinned(entry.Id, !entry.IsPinned);
                int idx = _history.FindIndex(h => h.Id == entry.Id);
                if (idx >= 0) _history[idx] = _history[idx] with { IsPinned = !entry.IsPinned };
                RefreshHistoryList();
            });

            // Rename / label
            string renameLabel = string.IsNullOrEmpty(entry.Label) ? "🏷 Set Name..." : $"🏷 Rename \"{entry.Label}\"...";
            menu.Items.Add(renameLabel, null, (s, _) =>
            {
                string? newLabel = SimplePrompt("Set Name", "Name (leave blank to clear):", entry.Label ?? "");
                if (newLabel == null) return;   // cancelled
                HistoryStore.SetLabel(entry.Id, newLabel);
                int idx = _history.FindIndex(h => h.Id == entry.Id);
                if (idx >= 0) _history[idx] = _history[idx] with
                {
                    Label = string.IsNullOrWhiteSpace(newLabel) ? null : newLabel.Trim()
                };
                RefreshHistoryList();
            });

            menu.Items.Add("Load in current tab", null, (s, _) =>
            {
                ActiveTab.Editor.Text = entry.Sql;
                ActiveTab.Editor.Focus();
            });

            menu.Items.Add("Open in new tab", null, (s, _) =>
            {
                var tab = CreateTab($"Query {tabQueries.TabPages.Count + 1}");
                tab.Editor.Text = entry.Sql;
                tabQueries.SelectedTab = tab;
                tab.Editor.Focus();
            });

            menu.Items.Add("Copy SQL", null, (s, _) => Clipboard.SetText(entry.Sql));

            menu.Items.Add("-");

            menu.Items.Add("Delete", null, (s, _) =>
            {
                HistoryStore.Delete(entry.Id);
                _history.RemoveAll(h => h.Id == entry.Id);
                RefreshHistoryList();
            });

            menu.Show(lvHistory, location);
        }

        private void txtHistoryFilter_TextChanged(object sender, EventArgs e)
        {
            _historyFilter = txtHistoryFilter.Text;
            RefreshHistoryList();
        }

        private void btnClearHistory_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Clear all unpinned history?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            HistoryStore.ClearUnpinned();
            _history.RemoveAll(h => !h.IsPinned);
            RefreshHistoryList();
        }

        // ── History ListView owner draw ──────────────────────────────────────────

        private void LvHistory_DrawColumnHeader(object? sender, DrawListViewColumnHeaderEventArgs e)
        {
            using var bg = new SolidBrush(Color.FromArgb(45, 45, 48));
            e.Graphics.FillRectangle(bg, e.Bounds);
            TextRenderer.DrawText(e.Graphics, e.Header!.Text, lvHistory.Font,
                e.Bounds, Color.FromArgb(150, 150, 150),
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }

        private void LvHistory_DrawItem(object? sender, DrawListViewItemEventArgs e)
        {
            // Drawing handled per-subitem
        }

        private void LvHistory_DrawSubItem(object? sender, DrawListViewSubItemEventArgs e)
        {
            if (e.Item == null) return;

            bool selected = e.Item.Selected;

            var bgColor = selected
                ? Color.FromArgb(0, 122, 204)
                : e.ItemIndex % 2 == 0
                    ? Color.FromArgb(30, 30, 30)
                    : Color.FromArgb(35, 35, 37);

            // Respect the per-item color set during refresh (gold=pinned, blue=labeled, normal)
            var fgColor = selected ? Color.White : e.Item.ForeColor;

            using var bgBrush = new SolidBrush(bgColor);
            e.Graphics.FillRectangle(bgBrush, e.Bounds);

            var textRect = new Rectangle(
                e.Bounds.X + 4, e.Bounds.Y,
                e.Bounds.Width - 4, e.Bounds.Height);

            var flags = TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis;
            if (e.ColumnIndex == 1) flags |= TextFormatFlags.Right;

            TextRenderer.DrawText(e.Graphics, e.SubItem?.Text ?? "",
                lvHistory.Font, textRect, fgColor, flags);
        }

        // ── Keyboard shortcut in editor ─────────────────────────────────────────

        private void rtbQuery_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                btnExecute_Click(sender, EventArgs.Empty);
            }
        }

        // ── Helpers ─────────────────────────────────────────────────────────────

        private void btnCompareDb_Click(object sender, EventArgs e)
        {
            using var dlg = new CompareDbDialog(_connectionString);
            dlg.ShowDialog(this);
        }

        private void btnDeployCheck_Click(object sender, EventArgs e)
        {
            using var dlg = new DeploymentCheckerDialog(_connectionString);
            dlg.ShowDialog(this);
        }

        private void UpdateConnectionStatus(bool connected)
        {
            picStatus.BackColor        = connected ? Color.LimeGreen : Color.Gray;
            btnDisconnect.Enabled      = connected;
            btnExecute.Enabled         = connected;
            btnExecuteNonQuery.Enabled = connected;
            btnCompareDb.Enabled       = connected;
            btnDeployCheck.Enabled     = connected;
        }

        private void SetExecutingState(bool executing)
        {
            btnExecute.Enabled         = !executing;
            btnExecuteNonQuery.Enabled = !executing;
            progressBar.Visible        = executing;
            lblStatus.Text             = executing ? "Executing..." : "Ready";
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
            var lbl = new Label
            {
                Text = label, Location = new Point(10, 12), AutoSize = true,
                ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9.5f)
            };
            var txt = new TextBox
            {
                Text = defaultValue, Location = new Point(10, 32), Width = 304,
                BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9.5f)
            };
            var ok = new Button
            {
                Text = "OK", Location = new Point(148, 64), Width = 75, Height = 28,
                DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White
            };
            ok.FlatAppearance.BorderSize = 0;
            var cancel = new Button
            {
                Text = "Cancel", Location = new Point(233, 64), Width = 75, Height = 28,
                DialogResult = DialogResult.Cancel, FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72), ForeColor = Color.White
            };
            cancel.FlatAppearance.BorderSize = 0;
            frm.AcceptButton = ok; frm.CancelButton = cancel;
            frm.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
            txt.SelectAll();
            return frm.ShowDialog() == DialogResult.OK ? txt.Text.Trim() : null;
        }
    }
}
