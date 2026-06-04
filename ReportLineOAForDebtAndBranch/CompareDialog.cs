using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ReportLineOAForDebtAndBranch
{
    public class CompareDialog : Form
    {
        private readonly DataTable _queryData;
        private DataTable? _excelData;
        private DataTable? _result;

        // Left panel (setup)
        private TextBox txtFile = null!;
        private Button btnBrowse = null!;
        private ComboBox cboSheet = null!;
        private FlowLayoutPanel pnlMappings = null!;
        private Button btnAddRow = null!;
        private Button btnCompare = null!;

        // Right panel (results)
        private DataGridView grid = null!;
        private Label lblSummary = null!;
        private Button btnExportCsv = null!;
        private Button btnExportXlsx = null!;

        private readonly List<MappingRow> _rows = new();

        private sealed class MappingRow
        {
            public Panel Container = null!;
            public RadioButton RdoKey = null!;
            public ComboBox QueryCol = null!;
            public ComboBox ExcelCol = null!;
        }

        public CompareDialog(DataTable queryData)
        {
            _queryData = queryData;
            BuildUI();
            AddRow();
        }

        // ── UI Construction ─────────────────────────────────────────────────────

        private void BuildUI()
        {
            Text          = "Compare with Excel";
            Size          = new Size(1150, 700);
            MinimumSize   = new Size(900, 550);
            StartPosition = FormStartPosition.CenterParent;
            BackColor     = Color.FromArgb(30, 30, 30);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            BuildSetupPanel(split.Panel1);
            BuildResultPanel(split.Panel2);
            Controls.Add(split);

            Load += (s, e) =>
            {
                split.Panel1MinSize = 330;
                split.Panel2MinSize = 380;
                split.SplitterDistance = 360;
            };
        }

        private void BuildSetupPanel(SplitterPanel pnl)
        {
            pnl.BackColor = Color.FromArgb(37, 37, 38);

            // ── Excel file ───────────────────────────────────────────────────────
            var grpFile = MakeGroup("Excel File", 8, 8);
            grpFile.Height = 82;

            txtFile = new TextBox
            {
                ReadOnly = true, Location = new Point(8, 22), Width = 210,
                BackColor = Color.FromArgb(55, 55, 55), ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle, Font = new Font("Segoe UI", 9f)
            };
            btnBrowse = MakeBtn("Browse...", 224, 20, 82);
            btnBrowse.Click += BtnBrowse_Click;

            var lblSheet = MakeLabel("Sheet:", 8, 52);
            cboSheet = new ComboBox
            {
                Location = new Point(54, 49), Width = 250,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(55, 55, 55), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f)
            };
            cboSheet.SelectedIndexChanged += CboSheet_Changed;
            grpFile.Controls.AddRange(new Control[] { txtFile, btnBrowse, lblSheet, cboSheet });

            // ── Column mapping ───────────────────────────────────────────────────
            var grpMap = MakeGroup("Column Mapping  (◉ = Key → match by value, ignores row order  |  no key → match by row position)", 8, 98);
            grpMap.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            var hRow = new Panel { Location = new Point(6, 18), Height = 18, BackColor = Color.Transparent };
            hRow.Controls.Add(MakeSmLabel("Key", 2, 0, 36));
            hRow.Controls.Add(MakeSmLabel("Query Column", 44, 0, 118));
            hRow.Controls.Add(MakeSmLabel("Excel Column", 172, 0, 118));
            grpMap.Controls.Add(hRow);

            pnlMappings = new FlowLayoutPanel
            {
                Location = new Point(6, 38),
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = Color.FromArgb(37, 37, 38)
            };
            grpMap.Controls.Add(pnlMappings);

            btnAddRow = MakeBtn("+ Add Column", 8, 0, 120);
            btnAddRow.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnAddRow.Click += (s, e) => AddRow();

            btnCompare = MakeBtn("▶  Compare", 0, 0, 110);
            btnCompare.BackColor = Color.FromArgb(0, 122, 204);
            btnCompare.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnCompare.Click += BtnCompare_Click;

            grpMap.Controls.AddRange(new Control[] { btnAddRow, btnCompare });

            // Resize handlers
            grpMap.Resize += (s, e) =>
            {
                int bY = grpMap.ClientSize.Height - 32;
                btnAddRow.Location  = new Point(8, bY);
                btnCompare.Location = new Point(grpMap.ClientSize.Width - 118, bY);
                pnlMappings.Width   = grpMap.ClientSize.Width - 14;
                pnlMappings.Height  = bY - 40;
                hRow.Width          = pnlMappings.Width;
                foreach (var r in _rows) r.Container.Width = pnlMappings.ClientSize.Width - 4;
            };

            pnl.Resize += (s, e) =>
            {
                int w = pnl.ClientSize.Width - 16;
                grpFile.Width = w;
                grpMap.Width  = w;
                grpMap.Height = pnl.ClientSize.Height - 110;
            };

            pnl.Controls.AddRange(new Control[] { grpFile, grpMap });
        }

        private void BuildResultPanel(SplitterPanel pnl)
        {
            pnl.BackColor = Color.FromArgb(30, 30, 30);

            lblSummary = new Label
            {
                Dock = DockStyle.Top, Height = 26,
                Text = "Run a comparison to see results",
                Font = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(150, 150, 150),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                BackColor = Color.FromArgb(45, 45, 48)
            };

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                BackgroundColor = Color.FromArgb(30, 30, 30), GridColor = Color.FromArgb(55, 55, 55),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 9f),
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
                EnableHeadersVisualStyles = false
            };
            grid.DefaultCellStyle.BackColor         = Color.FromArgb(30, 30, 30);
            grid.DefaultCellStyle.ForeColor         = Color.FromArgb(212, 212, 212);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 204);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(200, 200, 200);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(33, 33, 33);
            grid.RowPrePaint += Grid_RowPrePaint;
            grid.DataError   += (s, e) => e.Cancel = true;

            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom, Height = 38,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            btnExportCsv = MakeBtn("Export CSV", 6, 6, 100);
            btnExportCsv.Click += (s, e) => ExportCsv();

            btnExportXlsx = MakeBtn("Export Excel (.xlsx)", 112, 6, 148);
            btnExportXlsx.BackColor = Color.FromArgb(20, 100, 45);
            btnExportXlsx.Click += (s, e) => ExportExcel();

            pnlBottom.Controls.AddRange(new Control[] { btnExportCsv, btnExportXlsx });

            pnl.Controls.Add(grid);
            pnl.Controls.Add(pnlBottom);
            pnl.Controls.Add(lblSummary);
        }

        // ── Mapping Rows ────────────────────────────────────────────────────────

        private void AddRow()
        {
            var row = new MappingRow();

            row.Container = new Panel
            {
                Width  = Math.Max(pnlMappings.ClientSize.Width - 4, 300),
                Height = 30,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            row.RdoKey = new RadioButton
            {
                Location = new Point(4, 7), Width = 36,
                ForeColor = Color.FromArgb(180, 180, 180)
            };
            // Mutual exclusion across different parent panels
            row.RdoKey.CheckedChanged += (s, e) =>
            {
                if (!row.RdoKey.Checked) return;
                foreach (var other in _rows.Where(r => r != row))
                    other.RdoKey.Checked = false;
            };

            row.QueryCol = MakeCombo(42, 2, 118);
            foreach (DataColumn c in _queryData.Columns)
                row.QueryCol.Items.Add(c.ColumnName);
            if (_rows.Count < row.QueryCol.Items.Count)
                row.QueryCol.SelectedIndex = _rows.Count;
            else if (row.QueryCol.Items.Count > 0)
                row.QueryCol.SelectedIndex = 0;

            var lbl = new Label
            {
                Text = "↔", Location = new Point(164, 6), AutoSize = true,
                ForeColor = Color.FromArgb(100, 100, 100), Font = new Font("Segoe UI", 10f)
            };

            row.ExcelCol = MakeCombo(180, 2, 118);
            PopulateExcelCols(row.ExcelCol);

            var btnDel = MakeBtn("×", 302, 3, 26);
            btnDel.ForeColor = Color.FromArgb(200, 80, 80);
            btnDel.Click += (s, e) =>
            {
                if (_rows.Count <= 1) return;
                _rows.Remove(row);
                pnlMappings.Controls.Remove(row.Container);
            };

            row.Container.Controls.AddRange(new Control[]
            {
                row.RdoKey, row.QueryCol, lbl, row.ExcelCol, btnDel
            });

            _rows.Add(row);
            pnlMappings.Controls.Add(row.Container);

            if (_rows.Count == 1) row.RdoKey.Checked = true;
        }

        private void PopulateExcelCols(ComboBox cbo)
        {
            int prev = cbo.SelectedIndex;
            cbo.Items.Clear();
            if (_excelData == null) return;
            foreach (DataColumn c in _excelData.Columns) cbo.Items.Add(c.ColumnName);
            cbo.SelectedIndex = prev >= 0 && prev < cbo.Items.Count ? prev : cbo.Items.Count > 0 ? 0 : -1;
        }

        // ── Excel Loading ───────────────────────────────────────────────────────

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx;*.xls|All files|*.*",
                Title  = "Open Excel file"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            txtFile.Text = dlg.FileName;
            LoadSheets(dlg.FileName);
        }

        private void LoadSheets(string path)
        {
            try
            {
                cboSheet.Items.Clear();
                using var wb = new XLWorkbook(path);
                foreach (var ws in wb.Worksheets) cboSheet.Items.Add(ws.Name);
                if (cboSheet.Items.Count > 0) cboSheet.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading file:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CboSheet_Changed(object? sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFile.Text) || cboSheet.SelectedIndex < 0) return;
            try
            {
                using var wb = new XLWorkbook(txtFile.Text);
                var ws = wb.Worksheet(cboSheet.SelectedItem!.ToString()!);
                _excelData = SheetToDataTable(ws);
                foreach (var r in _rows) PopulateExcelCols(r.ExcelCol);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading sheet:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static DataTable SheetToDataTable(IXLWorksheet ws)
        {
            var dt   = new DataTable();
            var used = ws.RangeUsed();
            if (used == null) return dt;
            var rows = used.Rows().ToList();
            if (rows.Count == 0) return dt;

            foreach (var cell in rows[0].Cells())
                dt.Columns.Add(cell.Value.ToString());

            foreach (var row in rows.Skip(1))
            {
                var dr = dt.NewRow();
                int i  = 0;
                foreach (var cell in row.Cells())
                {
                    if (i < dt.Columns.Count) dr[i] = cell.Value.ToString();
                    i++;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        // ── Comparison Logic ────────────────────────────────────────────────────

        private void BtnCompare_Click(object? sender, EventArgs e)
        {
            if (_excelData == null)
            {
                MessageBox.Show("Please load an Excel file first.", "Missing Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var mappings = _rows
                .Where(r => r.QueryCol.SelectedItem != null && r.ExcelCol.SelectedItem != null)
                .Select(r => (
                    QCol:  r.QueryCol.SelectedItem!.ToString()!,
                    ECol:  r.ExcelCol.SelectedItem!.ToString()!,
                    IsKey: r.RdoKey.Checked))
                .ToList();

            if (mappings.Count == 0)
            {
                MessageBox.Show("Add at least one column mapping.", "Missing Mapping",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                _result = RunCompare(mappings);
                grid.DataSource = _result;
                AdjustColumns();
                UpdateSummary();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Compare error:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable RunCompare(List<(string QCol, string ECol, bool IsKey)> mappings)
        {
            var dt     = new DataTable();
            var keyMap = mappings.FirstOrDefault(m => m.IsKey);
            bool hasKey = !string.IsNullOrEmpty(keyMap.QCol);
            var nonKey  = mappings.Where(m => !m.IsKey).ToList();

            dt.Columns.Add("Status");
            if (hasKey) dt.Columns.Add(keyMap.QCol + " (Key)");
            foreach (var m in nonKey)
            {
                dt.Columns.Add(m.QCol + " (Query)");
                dt.Columns.Add(m.ECol + " (Excel)");
                dt.Columns.Add("✓ " + m.QCol);
            }

            if (hasKey)
            {
                // Build Excel lookup
                var lookup = new Dictionary<string, DataRow>(StringComparer.OrdinalIgnoreCase);
                foreach (DataRow er in _excelData!.Rows)
                {
                    string k = er[keyMap.ECol]?.ToString()?.Trim() ?? "";
                    if (!lookup.ContainsKey(k)) lookup[k] = er;
                }

                // Query is master — iterate all query rows, look up Excel by key value (order-independent)
                foreach (DataRow qr in _queryData.Rows)
                {
                    string key = qr[keyMap.QCol]?.ToString()?.Trim() ?? "";
                    var row = dt.NewRow();
                    row[keyMap.QCol + " (Key)"] = key;

                    if (lookup.TryGetValue(key, out var er))
                    {
                        bool allMatch = true;
                        foreach (var m in nonKey)
                        {
                            string qv = qr[m.QCol]?.ToString() ?? "";
                            string ev = er[m.ECol]?.ToString()  ?? "";
                            bool   ok = string.Equals(qv.Trim(), ev.Trim(), StringComparison.OrdinalIgnoreCase);
                            row[m.QCol + " (Query)"] = qv;
                            row[m.ECol + " (Excel)"] = ev;
                            row["✓ " + m.QCol]       = ok ? "✅" : "❌";
                            if (!ok) allMatch = false;
                        }
                        row["Status"] = allMatch ? "✅ Match" : "❌ Mismatch";
                    }
                    else
                    {
                        // Query row has no matching row in Excel
                        row["Status"] = "⚠ Not in Excel";
                        foreach (var m in nonKey)
                            row[m.QCol + " (Query)"] = qr[m.QCol]?.ToString() ?? "";
                    }
                    dt.Rows.Add(row);
                }
                // Excel rows that have no matching query row are NOT shown — Query is the master
            }
            else
            {
                // No key selected → row-order comparison, Query is still master
                // Only iterate query rows; Excel rows beyond query count are ignored
                for (int i = 0; i < _queryData.Rows.Count; i++)
                {
                    var row = dt.NewRow();

                    if (i >= _excelData!.Rows.Count)
                    {
                        row["Status"] = "⚠ Not in Excel";
                        foreach (var m in mappings)
                            row[m.QCol + " (Query)"] = _queryData.Rows[i][m.QCol]?.ToString() ?? "";
                        dt.Rows.Add(row);
                        continue;
                    }

                    var qr = _queryData.Rows[i];
                    var er = _excelData.Rows[i];
                    bool allMatch = true;

                    foreach (var m in mappings)
                    {
                        string qv = qr[m.QCol]?.ToString() ?? "";
                        string ev = er[m.ECol]?.ToString()  ?? "";
                        bool   ok = string.Equals(qv.Trim(), ev.Trim(), StringComparison.OrdinalIgnoreCase);
                        row[m.QCol + " (Query)"] = qv;
                        row[m.ECol + " (Excel)"] = ev;
                        row["✓ " + m.QCol]       = ok ? "✅" : "❌";
                        if (!ok) allMatch = false;
                    }
                    row["Status"] = allMatch ? "✅ Match" : "❌ Mismatch";
                    dt.Rows.Add(row);
                }
                // Excel rows beyond query count are NOT shown — Query is the master
            }

            return dt;
        }

        // ── Grid Appearance ─────────────────────────────────────────────────────

        private void AdjustColumns()
        {
            if (grid.Columns.Contains("Status"))
                grid.Columns["Status"]!.Width = 120;

            foreach (DataGridViewColumn col in grid.Columns)
                if (col.Name.StartsWith("✓ ")) col.Width = 38;
        }

        private void Grid_RowPrePaint(object? sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= grid.Rows.Count) return;
            if (grid.Rows[e.RowIndex].DataBoundItem is not DataRowView drv) return;

            string status = drv.Row["Status"]?.ToString() ?? "";
            var bg = status switch
            {
                "✅ Match"     => Color.FromArgb(18, 50, 22),
                "❌ Mismatch"  => Color.FromArgb(62, 18, 18),
                _             => Color.FromArgb(55, 44, 18)
            };
            grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = bg;
        }

        private void UpdateSummary()
        {
            if (_result == null) return;
            int total       = _result.Rows.Count;
            int match       = _result.AsEnumerable().Count(r => r["Status"].ToString() == "✅ Match");
            int mismatch    = _result.AsEnumerable().Count(r => r["Status"].ToString() == "❌ Mismatch");
            int notInExcel  = total - match - mismatch;

            lblSummary.Text =
                $"Query rows: {total}    " +
                $"✅ Match: {match}    " +
                $"❌ Mismatch: {mismatch}    " +
                $"⚠ Not in Excel: {notInExcel}    " +
                "(Query = master — Excel rows with no matching Query row are excluded)";

            lblSummary.ForeColor = mismatch > 0
                ? Color.FromArgb(255, 130, 130)
                : Color.FromArgb(100, 220, 100);
        }

        // ── Export ──────────────────────────────────────────────────────────────

        private void ExportCsv()
        {
            if (_result == null || _result.Rows.Count == 0) return;
            using var dlg = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "compare_result.csv" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", _result.Columns.Cast<DataColumn>()
                .Select(c => Esc(c.ColumnName))));
            foreach (DataRow row in _result.Rows)
                sb.AppendLine(string.Join(",", row.ItemArray.Select(v => Esc(v?.ToString() ?? ""))));

            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            MessageBox.Show("Exported successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportExcel()
        {
            if (_result == null || _result.Rows.Count == 0) return;
            using var dlg = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "compare_result.xlsx" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Compare Result");

            // Headers
            for (int c = 0; c < _result.Columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = _result.Columns[c].ColumnName;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(37, 37, 38);
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Font.Bold = true;
            }

            // Data rows with color coding
            for (int r = 0; r < _result.Rows.Count; r++)
            {
                string status = _result.Rows[r]["Status"]?.ToString() ?? "";
                var rowBg = status switch
                {
                    "✅ Match"    => XLColor.FromArgb(198, 239, 206),
                    "❌ Mismatch" => XLColor.FromArgb(255, 199, 206),
                    _            => XLColor.FromArgb(255, 235, 156)
                };

                for (int c = 0; c < _result.Columns.Count; c++)
                {
                    var cell = ws.Cell(r + 2, c + 1);
                    cell.Value = _result.Rows[r][c]?.ToString() ?? "";
                    cell.Style.Fill.BackgroundColor = rowBg;

                    // Extra: highlight mismatch check columns
                    string colName = _result.Columns[c].ColumnName;
                    if (colName.StartsWith("✓ ") && cell.Value.ToString() == "❌")
                        cell.Style.Font.FontColor = XLColor.Red;
                }
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
            wb.SaveAs(dlg.FileName);
            MessageBox.Show("Exported successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Helper factory methods ───────────────────────────────────────────────

        private static GroupBox MakeGroup(string title, int x, int y)
            => new GroupBox
            {
                Text = title, Location = new Point(x, y),
                ForeColor = Color.FromArgb(150, 150, 150),
                Font = new Font("Segoe UI", 8.5f)
            };

        private static Button MakeBtn(string text, int x, int y, int w)
        {
            var b = new Button
            {
                Text = text, Location = new Point(x, y), Width = w, Height = 27,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(65, 65, 68), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private static ComboBox MakeCombo(int x, int y, int w)
            => new ComboBox
            {
                Location = new Point(x, y), Width = w,
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(55, 55, 55), ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f)
            };

        private static Label MakeLabel(string text, int x, int y)
            => new Label
            {
                Text = text, Location = new Point(x, y), AutoSize = true,
                ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 9f)
            };

        private static Label MakeSmLabel(string text, int x, int y, int w)
            => new Label
            {
                Text = text, Location = new Point(x, y), Width = w, AutoSize = false,
                ForeColor = Color.FromArgb(120, 120, 120), Font = new Font("Segoe UI", 8f, FontStyle.Italic)
            };

        private static string Esc(string s) =>
            s.Contains(',') || s.Contains('"') || s.Contains('\n')
                ? $"\"{s.Replace("\"", "\"\"")}\"" : s;
    }
}
