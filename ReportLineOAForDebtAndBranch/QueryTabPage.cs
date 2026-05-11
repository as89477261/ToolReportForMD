using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace ReportLineOAForDebtAndBranch
{
    public class QueryTabPage : TabPage
    {
        public RichTextBox Editor { get; private set; } = null!;
        public DataGridView Grid { get; private set; } = null!;
        public RichTextBox Messages { get; private set; } = null!;
        public TabControl ResultTabs { get; private set; } = null!;
        public TabPage TabResults { get; private set; } = null!;
        public TabPage TabMessages { get; private set; } = null!;
        public Label LblRowCount { get; private set; } = null!;

        private static int _counter = 1;

        public QueryTabPage() : base($"Query {_counter++}") => BuildUI();

        private void BuildUI()
        {
            BackColor = Color.FromArgb(30, 30, 30);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                Panel1MinSize = 60,
                Panel2MinSize = 60,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            // ── Query editor ─────────────────────────────────────────────────────
            Editor = new RichTextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 11f),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                AcceptsTab = true,
                ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false
            };
            split.Panel1.Controls.Add(Editor);

            // ── Results toolbar ──────────────────────────────────────────────────
            var pnlResultsToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 34,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            var btnClear = MakeBtn("Clear", Color.FromArgb(70, 70, 72), 4, 4, 60);
            btnClear.Click += (s, e) => ClearResults();

            var btnExport = MakeBtn("Export CSV", Color.FromArgb(70, 70, 72), 70, 4, 90);
            btnExport.Click += (s, e) => ExportCsv();

            LblRowCount = new Label
            {
                AutoSize = true,
                Location = new Point(170, 9),
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(160, 160, 160)
            };

            pnlResultsToolbar.Controls.AddRange(new Control[] { btnClear, btnExport, LblRowCount });

            // ── DataGridView ─────────────────────────────────────────────────────
            Grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                BackgroundColor = Color.FromArgb(30, 30, 30),
                GridColor = Color.FromArgb(60, 60, 60),
                BorderStyle = BorderStyle.None,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersWidth = 30,
                Font = new Font("Consolas", 9.5f)
            };
            Grid.DefaultCellStyle.BackColor = Color.FromArgb(30, 30, 30);
            Grid.DefaultCellStyle.ForeColor = Color.FromArgb(212, 212, 212);
            Grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 204);
            Grid.DefaultCellStyle.SelectionForeColor = Color.White;
            Grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            Grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(200, 200, 200);
            Grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            Grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            Grid.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            Grid.EnableHeadersVisualStyles = false;
            // Suppress errors for unsupported column types (binary/image)
            Grid.DataError += (s, e) => e.Cancel = true;

            // ── Messages ─────────────────────────────────────────────────────────
            Messages = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10f),
                BackColor = Color.FromArgb(20, 20, 20),
                ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            TabResults = new TabPage("Results");
            TabResults.Controls.Add(Grid);

            TabMessages = new TabPage("Messages");
            TabMessages.Controls.Add(Messages);

            ResultTabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f)
            };
            ResultTabs.TabPages.AddRange(new[] { TabResults, TabMessages });

            split.Panel2.Controls.Add(ResultTabs);
            split.Panel2.Controls.Add(pnlResultsToolbar);

            Controls.Add(split);
        }

        private static Button MakeBtn(string text, Color bg, int x, int y, int w)
        {
            var btn = new Button
            {
                Text = text, Location = new Point(x, y), Width = w, Height = 26,
                FlatStyle = FlatStyle.Flat, BackColor = bg, ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        public string GetActiveQuery()
        {
            string sel = Editor.SelectedText.Trim();
            return string.IsNullOrEmpty(sel) ? Editor.Text.Trim() : sel;
        }

        public void SetResult(DataTable? dt, long elapsedMs)
        {
            Grid.DataSource = dt;
            int count = dt?.Rows.Count ?? 0;
            LblRowCount.Text = $"{count} row(s)  |  {elapsedMs} ms";
        }

        public void AppendMessage(string text, Color color)
        {
            if (Messages.InvokeRequired) { Messages.Invoke(() => AppendMessage(text, color)); return; }
            int start = Messages.TextLength;
            string line = $"[{DateTime.Now:HH:mm:ss}]  {text}{Environment.NewLine}";
            Messages.AppendText(line);
            Messages.Select(start, line.Length);
            Messages.SelectionColor = color;
            Messages.SelectionLength = 0;
            Messages.ScrollToCaret();
        }

        public void ClearResults()
        {
            Grid.DataSource = null;
            Messages.Clear();
            LblRowCount.Text = string.Empty;
        }

        public void ShowResultsTab() => ResultTabs.SelectedTab = TabResults;
        public void ShowMessagesTab() => ResultTabs.SelectedTab = TabMessages;

        private void ExportCsv()
        {
            if (Grid.DataSource is not DataTable dt || dt.Rows.Count == 0) return;
            using var dlg = new SaveFileDialog { Filter = "CSV files|*.csv", FileName = "result.csv" };
            if (dlg.ShowDialog() != DialogResult.OK) return;
            try
            {
                var lines = new List<string>();
                var headers = new List<string>();
                foreach (DataColumn col in dt.Columns) headers.Add(CsvEscape(col.ColumnName));
                lines.Add(string.Join(",", headers));
                foreach (DataRow row in dt.Rows)
                {
                    var cells = new List<string>();
                    foreach (var item in row.ItemArray) cells.Add(CsvEscape(item?.ToString() ?? ""));
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
