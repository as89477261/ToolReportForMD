using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;

namespace ReportLineOAForDebtAndBranch
{
    public class CompareDbDialog : Form
    {
        private readonly string _connectionString;

        // Mode
        private RadioButton rdoData      = null!;
        private RadioButton rdoStructure = null!;
        private RadioButton rdoProc      = null!;

        // Source / Target selectors
        private ComboBox cboLeftDb  = null!, cboLeftSchema  = null!, cboLeftObj  = null!;
        private ComboBox cboRightDb = null!, cboRightSchema = null!, cboRightObj = null!;

        // Key panel (data mode)
        private Panel    pnlKey     = null!;
        private ComboBox cboKeyLeft = null!, cboKeyRight = null!;

        // Compare button
        private Button btnCompare = null!;

        // Results
        private Label       lblSummary   = null!;
        private DataGridView grid        = null!;
        private RichTextBox  txtDiff     = null!;
        private Button btnExportCsv      = null!;
        private Button btnExportXlsx     = null!;

        private DataTable? _result;

        public CompareDbDialog(string connectionString)
        {
            _connectionString = connectionString;
            BuildUI();
        }

        // ── UI ──────────────────────────────────────────────────────────────────

        private void BuildUI()
        {
            Text          = "Compare Database Objects";
            Size          = new Size(1160, 800);
            MinimumSize   = new Size(900, 640);
            StartPosition = FormStartPosition.CenterParent;
            BackColor     = Color.FromArgb(30, 30, 30);

            // Mode bar
            var pnlMode = new Panel
            {
                Dock = DockStyle.Top, Height = 40,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            rdoData      = MakeRdo("Table Data",        80,  10);
            rdoStructure = MakeRdo("Table Structure",   210, 10);
            rdoProc      = MakeRdo("Stored Procedure",  360, 10);
            rdoData.Checked = true;

            rdoData.CheckedChanged      += (s, e) => { if (rdoData.Checked)      ModeChanged(); };
            rdoStructure.CheckedChanged += (s, e) => { if (rdoStructure.Checked) ModeChanged(); };
            rdoProc.CheckedChanged      += (s, e) => { if (rdoProc.Checked)      ModeChanged(); };

            pnlMode.Controls.AddRange(new Control[]
            {
                MakeLabel("Compare:", 10, 12), rdoData, rdoStructure, rdoProc
            });

            // Sources row
            var pnlSources = new Panel
            {
                Dock = DockStyle.Top, Height = 156,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            var grpLeft  = MakeGroup("Source", 8, 8);
            BuildSourcePanel(grpLeft,  out cboLeftDb,  out cboLeftSchema,  out cboLeftObj,  isLeft: true);

            var grpRight = MakeGroup("Target", 0, 8);
            BuildSourcePanel(grpRight, out cboRightDb, out cboRightSchema, out cboRightObj, isLeft: false);

            // Key panel
            pnlKey = new Panel { Location = new Point(0, 8), BackColor = Color.FromArgb(37, 37, 38) };
            var grpKey = MakeGroup("Key Columns (for row matching — leave blank = row-position mode)", 0, 0);
            grpKey.Dock = DockStyle.Fill;

            cboKeyLeft  = MakeCombo(8, 36, 200);
            cboKeyRight = MakeCombo(8, 78, 200);
            grpKey.Controls.Add(MakeSmLabel("Source key column:", 8, 20, 180));
            grpKey.Controls.Add(cboKeyLeft);
            grpKey.Controls.Add(MakeSmLabel("Target key column:", 8, 62, 180));
            grpKey.Controls.Add(cboKeyRight);
            grpKey.Controls.Add(MakeSmLabel("Without a key, rows are matched by position", 8, 112, 230));
            pnlKey.Controls.Add(grpKey);

            pnlSources.Controls.AddRange(new Control[] { grpLeft, grpRight, pnlKey });
            pnlSources.Resize += (s, e) =>
            {
                int w = (int)((pnlSources.ClientSize.Width - 32) * 0.38);
                grpLeft.Width  = w; grpLeft.Height  = pnlSources.ClientSize.Height - 16;
                grpRight.Width = w; grpRight.Height = grpLeft.Height;
                grpRight.Location = new Point(grpLeft.Right + 8, 8);
                pnlKey.Location   = new Point(grpRight.Right + 8, 8);
                pnlKey.Width      = pnlSources.ClientSize.Width - pnlKey.Left - 8;
                pnlKey.Height     = grpLeft.Height;
                foreach (var cb in new[] { cboKeyLeft, cboKeyRight })
                    cb.Width = pnlKey.ClientSize.Width - 20;
                foreach (Control c in grpLeft.Controls)
                    if (c is ComboBox cb2) cb2.Width = grpLeft.ClientSize.Width - 98;
                foreach (Control c in grpRight.Controls)
                    if (c is ComboBox cb2) cb2.Width = grpRight.ClientSize.Width - 98;
            };

            // Action bar
            var pnlAction = new Panel
            {
                Dock = DockStyle.Top, Height = 38,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            btnCompare = MakeBtn("▶  Compare", 0, 5, 140);
            btnCompare.BackColor = Color.FromArgb(0, 122, 204);
            btnCompare.Anchor    = AnchorStyles.Right;
            btnCompare.Click    += BtnCompare_Click;
            pnlAction.Controls.Add(btnCompare);
            pnlAction.Resize += (s, e) => btnCompare.Location = new Point(pnlAction.Width - 152, 5);

            // Summary bar
            lblSummary = new Label
            {
                Dock = DockStyle.Top, Height = 28,
                Text = "Select source and target objects, then click Compare.",
                Font = new Font("Segoe UI", 9f), ForeColor = Color.FromArgb(150, 150, 150),
                TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(6, 0, 0, 0),
                BackColor = Color.FromArgb(45, 45, 48)
            };

            // DataGridView
            grid = new DataGridView
            {
                Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false,
                AllowUserToDeleteRows = false, BackgroundColor = Color.FromArgb(30, 30, 30),
                GridColor = Color.FromArgb(55, 55, 55), BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 9f), EnableHeadersVisualStyles = false
            };
            grid.DefaultCellStyle.BackColor          = Color.FromArgb(30, 30, 30);
            grid.DefaultCellStyle.ForeColor          = Color.FromArgb(212, 212, 212);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 204);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(200, 200, 200);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(33, 33, 33);
            grid.RowPrePaint += Grid_RowPrePaint;
            grid.DataError   += (s, e) => e.Cancel = true;

            // Diff viewer (stored proc mode)
            txtDiff = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true, Visible = false,
                Font = new Font("Consolas", 10f),
                BackColor = Color.FromArgb(20, 20, 20), ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None, ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false
            };

            // Bottom bar
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom, Height = 38,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            btnExportCsv  = MakeBtn("Export CSV",         6,   6, 100);
            btnExportXlsx = MakeBtn("Export Excel (.xlsx)", 112, 6, 148);
            btnExportXlsx.BackColor = Color.FromArgb(20, 100, 45);
            btnExportCsv.Click  += (s, e) => ExportCsv();
            btnExportXlsx.Click += (s, e) => ExportExcel();
            pnlBottom.Controls.AddRange(new Control[] { btnExportCsv, btnExportXlsx });

            // Assemble
            var pnlContent = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(30, 30, 30) };
            pnlContent.Controls.Add(grid);
            pnlContent.Controls.Add(txtDiff);
            pnlContent.Controls.Add(pnlBottom);
            pnlContent.Controls.Add(lblSummary);

            SuspendLayout();
            Controls.Add(pnlContent);
            Controls.Add(pnlAction);
            Controls.Add(pnlSources);
            Controls.Add(pnlMode);
            ResumeLayout();

            Load += async (s, e) => await LoadDatabasesAsync();
        }

        private void BuildSourcePanel(GroupBox grp, out ComboBox cbDb, out ComboBox cbSchema,
            out ComboBox cbObj, bool isLeft)
        {
            cbDb     = MakeCombo(84, 19, 200);
            cbSchema = MakeCombo(84, 49, 200);
            cbObj    = MakeCombo(84, 79, 200);

            grp.Height = 140;
            grp.Controls.Add(MakeSmLabel("Database:", 8, 23, 74)); grp.Controls.Add(cbDb);
            grp.Controls.Add(MakeSmLabel("Schema:",   8, 53, 74)); grp.Controls.Add(cbSchema);
            grp.Controls.Add(MakeSmLabel("Object:",   8, 83, 74)); grp.Controls.Add(cbObj);

            // Closures
            var myDb     = cbDb;
            var mySch    = cbSchema;
            var myObj    = cbObj;
            var myIsLeft = isLeft;

            cbDb.SelectedIndexChanged += async (s, e) =>
            {
                if (myDb.SelectedItem == null) return;
                await LoadSchemasAsync(myDb.SelectedItem.ToString()!, mySch);
            };
            cbSchema.SelectedIndexChanged += async (s, e) =>
            {
                if (myDb.SelectedItem == null || mySch.SelectedItem == null) return;
                await LoadObjectsAsync(myDb.SelectedItem.ToString()!, mySch.SelectedItem.ToString()!, myObj);
                if (myIsLeft) await PopulateKeyColumnsAsync();
            };
            cbObj.SelectedIndexChanged += async (s, e) =>
            {
                if (myIsLeft) await PopulateKeyColumnsAsync();
            };
        }

        // ── Database / Object Loading ────────────────────────────────────────────

        private async Task LoadDatabasesAsync()
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                await cn.OpenAsync();
                var dt = new DataTable();
                using var da = new SqlDataAdapter(
                    "SELECT name FROM sys.databases WHERE state_desc='ONLINE' ORDER BY name", cn);
                await Task.Run(() => da.Fill(dt));

                var names   = dt.AsEnumerable().Select(r => r[0].ToString()!).ToList();
                string curDb = cn.Database;

                foreach (var cbo in new[] { cboLeftDb, cboRightDb })
                {
                    cbo.Items.Clear();
                    names.ForEach(n => cbo.Items.Add(n));
                    int idx = names.IndexOf(curDb);
                    cbo.SelectedIndex = idx >= 0 ? idx : cbo.Items.Count > 0 ? 0 : -1;
                }
            }
            catch (Exception ex) { lblSummary.Text = $"Error loading databases: {ex.Message}"; }
        }

        private async Task LoadSchemasAsync(string db, ComboBox cbSchema)
        {
            try
            {
                using var cn = new SqlConnection(_connectionString);
                await cn.OpenAsync();
                using var cmd = new SqlCommand(
                    $"SELECT SCHEMA_NAME FROM [{db}].INFORMATION_SCHEMA.SCHEMATA ORDER BY SCHEMA_NAME", cn);
                using var rd = await cmd.ExecuteReaderAsync();
                var list = new List<string>();
                while (await rd.ReadAsync()) list.Add(rd.GetString(0));
                cbSchema.Items.Clear();
                list.ForEach(s => cbSchema.Items.Add(s));
                int dbo = list.IndexOf("dbo");
                cbSchema.SelectedIndex = dbo >= 0 ? dbo : list.Count > 0 ? 0 : -1;
            }
            catch { }
        }

        private async Task LoadObjectsAsync(string db, string schema, ComboBox cbObj)
        {
            try
            {
                string sch = schema.Replace("'", "''");
                string sql = rdoProc.Checked
                    ? $"SELECT ROUTINE_NAME FROM [{db}].INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA='{sch}' AND ROUTINE_TYPE='PROCEDURE' ORDER BY ROUTINE_NAME"
                    : $"SELECT TABLE_NAME FROM [{db}].INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA='{sch}' AND TABLE_TYPE='BASE TABLE' ORDER BY TABLE_NAME";
                using var cn = new SqlConnection(_connectionString);
                await cn.OpenAsync();
                using var cmd = new SqlCommand(sql, cn);
                using var rd  = await cmd.ExecuteReaderAsync();
                var list = new List<string>();
                while (await rd.ReadAsync()) list.Add(rd.GetString(0));
                cbObj.Items.Clear();
                list.ForEach(o => cbObj.Items.Add(o));
                if (list.Count > 0) cbObj.SelectedIndex = 0;
            }
            catch { }
        }

        private async Task PopulateKeyColumnsAsync()
        {
            if (!rdoData.Checked || cboLeftDb.SelectedItem == null ||
                cboLeftSchema.SelectedItem == null || cboLeftObj.SelectedItem == null) return;
            try
            {
                string db  = cboLeftDb.SelectedItem.ToString()!;
                string sch = cboLeftSchema.SelectedItem.ToString()!.Replace("'", "''");
                string tbl = cboLeftObj.SelectedItem.ToString()!.Replace("'", "''");
                using var cn = new SqlConnection(_connectionString);
                await cn.OpenAsync();
                using var cmd = new SqlCommand(
                    $"SELECT COLUMN_NAME FROM [{db}].INFORMATION_SCHEMA.COLUMNS " +
                    $"WHERE TABLE_SCHEMA='{sch}' AND TABLE_NAME='{tbl}' ORDER BY ORDINAL_POSITION", cn);
                using var rd  = await cmd.ExecuteReaderAsync();
                var cols = new List<string> { "(none — positional)" };
                while (await rd.ReadAsync()) cols.Add(rd.GetString(0));
                cboKeyLeft.Items.Clear();  cboKeyRight.Items.Clear();
                cols.ForEach(c => { cboKeyLeft.Items.Add(c); cboKeyRight.Items.Add(c); });
                cboKeyLeft.SelectedIndex = cboKeyRight.SelectedIndex = 0;
            }
            catch { }
        }

        // ── Mode Change ──────────────────────────────────────────────────────────

        private void ModeChanged()
        {
            pnlKey.Visible = rdoData.Checked;
            ReloadObjectsForMode();
        }

        private async void ReloadObjectsForMode()
        {
            var tasks = new List<Task>();
            if (cboLeftDb.SelectedItem != null && cboLeftSchema.SelectedItem != null)
                tasks.Add(LoadObjectsAsync(cboLeftDb.SelectedItem.ToString()!, cboLeftSchema.SelectedItem.ToString()!, cboLeftObj));
            if (cboRightDb.SelectedItem != null && cboRightSchema.SelectedItem != null)
                tasks.Add(LoadObjectsAsync(cboRightDb.SelectedItem.ToString()!, cboRightSchema.SelectedItem.ToString()!, cboRightObj));
            await Task.WhenAll(tasks);
        }

        // ── Compare Dispatch ─────────────────────────────────────────────────────

        private async void BtnCompare_Click(object? sender, EventArgs e)
        {
            if (!ValidateSelections()) return;
            btnCompare.Enabled = false;
            lblSummary.Text = "Comparing…";
            lblSummary.ForeColor = Color.FromArgb(150, 150, 150);
            try
            {
                if      (rdoData.Checked)      await CompareDataAsync();
                else if (rdoStructure.Checked) await CompareStructureAsync();
                else                           await CompareProcAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Compare error:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblSummary.Text = "Error — see message box.";
            }
            finally { btnCompare.Enabled = true; }
        }

        private bool ValidateSelections()
        {
            bool ok = cboLeftDb.SelectedItem  != null && cboLeftSchema.SelectedItem  != null && cboLeftObj.SelectedItem  != null &&
                      cboRightDb.SelectedItem != null && cboRightSchema.SelectedItem != null && cboRightObj.SelectedItem != null;
            if (!ok) MessageBox.Show("Please select source and target objects.", "Missing Selection",
                         MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return ok;
        }

        private (string db, string sch, string obj) Left  =>
            (cboLeftDb.SelectedItem!.ToString()!, cboLeftSchema.SelectedItem!.ToString()!, cboLeftObj.SelectedItem!.ToString()!);
        private (string db, string sch, string obj) Right =>
            (cboRightDb.SelectedItem!.ToString()!, cboRightSchema.SelectedItem!.ToString()!, cboRightObj.SelectedItem!.ToString()!);

        // ── Table Data Compare ───────────────────────────────────────────────────

        private async Task CompareDataAsync()
        {
            var (lDb, lSch, lObj) = Left;
            var (rDb, rSch, rObj) = Right;
            string keyL = KeyColName(cboKeyLeft);
            string keyR = KeyColName(cboKeyRight);

            using var cn = new SqlConnection(_connectionString);
            await cn.OpenAsync();
            var leftDt  = await FetchTableAsync(cn, lDb, lSch, lObj);
            var rightDt = await FetchTableAsync(cn, rDb, rSch, rObj);

            _result = BuildDataCompare(leftDt, rightDt, keyL, keyR, lObj, rObj);
            ShowGrid();

            int total    = _result.Rows.Count;
            int match    = CountStatus("✅ Match");
            int mismatch = CountStatus("❌ Mismatch");
            int notFound = total - match - mismatch;
            lblSummary.Text =
                $"Table Data: [{lDb}].{lObj}  vs  [{rDb}].{rObj}   " +
                $"Total: {total}   ✅ Match: {match}   ❌ Mismatch: {mismatch}   ⚠ Not Found: {notFound}";
            lblSummary.ForeColor = mismatch + notFound > 0
                ? Color.FromArgb(255, 130, 130) : Color.FromArgb(100, 220, 100);
        }

        private static string KeyColName(ComboBox cbo)
        {
            string? s = cbo.SelectedItem?.ToString();
            return (s == null || s.StartsWith("(")) ? "" : s;
        }

        private static async Task<DataTable> FetchTableAsync(SqlConnection cn, string db, string sch, string tbl)
        {
            var dt  = new DataTable();
            using var da = new SqlDataAdapter($"SELECT * FROM [{db}].[{sch}].[{tbl}]", cn);
            await Task.Run(() => da.Fill(dt));
            return dt;
        }

        private static DataTable BuildDataCompare(DataTable left, DataTable right,
            string keyL, string keyR, string leftName, string rightName)
        {
            bool hasKey = !string.IsNullOrEmpty(keyL) && !string.IsNullOrEmpty(keyR);

            var lCols  = left.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var rCols  = right.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
            var common = lCols.Intersect(rCols, StringComparer.OrdinalIgnoreCase)
                              .Where(c => !c.Equals(keyL, StringComparison.OrdinalIgnoreCase))
                              .ToList();
            var onlyL  = lCols.Except(rCols, StringComparer.OrdinalIgnoreCase)
                               .Except(new[] { keyL }, StringComparer.OrdinalIgnoreCase).ToList();
            var onlyR  = rCols.Except(lCols, StringComparer.OrdinalIgnoreCase)
                               .Except(new[] { keyR }, StringComparer.OrdinalIgnoreCase).ToList();

            var dt = new DataTable();
            dt.Columns.Add("Status");
            if (hasKey) dt.Columns.Add($"{keyL} (Key)");
            foreach (var c in common)
            {
                dt.Columns.Add($"{c} ({leftName})");
                dt.Columns.Add($"{c} ({rightName})");
                dt.Columns.Add($"✓ {c}");
            }
            foreach (var c in onlyL)  dt.Columns.Add($"{c} (Source only)");
            foreach (var c in onlyR)  dt.Columns.Add($"{c} (Target only)");

            if (hasKey)
            {
                var rLookup = BuildLookup(right, keyR);
                var lLookup = BuildLookup(left,  keyL);

                foreach (DataRow lr in left.Rows)
                {
                    string k = Val(lr, keyL);
                    var row  = dt.NewRow();
                    row[$"{keyL} (Key)"] = k;
                    foreach (var c in onlyL) row[$"{c} (Source only)"] = Val(lr, c);
                    if (rLookup.TryGetValue(k, out var rr))
                    {
                        bool allOk = true;
                        foreach (var c in common)
                        {
                            string lv = Val(lr, c), rv = Val(rr, c);
                            bool   ok = string.Equals(lv.Trim(), rv.Trim(), StringComparison.OrdinalIgnoreCase);
                            row[$"{c} ({leftName})"]  = lv;
                            row[$"{c} ({rightName})"] = rv;
                            row[$"✓ {c}"]             = ok ? "✅" : "❌";
                            if (!ok) allOk = false;
                        }
                        foreach (var c in onlyR) row[$"{c} (Target only)"] = Val(rr, c);
                        row["Status"] = allOk ? "✅ Match" : "❌ Mismatch";
                    }
                    else
                    {
                        row["Status"] = "⚠ Not in Target";
                        foreach (var c in common) row[$"{c} ({leftName})"] = Val(lr, c);
                    }
                    dt.Rows.Add(row);
                }
                // Rows in right not in left
                foreach (DataRow rr in right.Rows)
                {
                    string k = Val(rr, keyR);
                    if (lLookup.ContainsKey(k)) continue;
                    var row = dt.NewRow();
                    row[$"{keyL} (Key)"] = k;
                    row["Status"] = "⚠ Not in Source";
                    foreach (var c in common) row[$"{c} ({rightName})"] = Val(rr, c);
                    foreach (var c in onlyR)  row[$"{c} (Target only)"] = Val(rr, c);
                    dt.Rows.Add(row);
                }
            }
            else
            {
                int max = Math.Max(left.Rows.Count, right.Rows.Count);
                for (int i = 0; i < max; i++)
                {
                    var row  = dt.NewRow();
                    bool hasL2 = i < left.Rows.Count;
                    bool hasR2 = i < right.Rows.Count;
                    if (!hasL2)
                    {
                        row["Status"] = "⚠ Not in Source";
                        foreach (var c in common) row[$"{c} ({rightName})"] = Val(right.Rows[i], c);
                        foreach (var c in onlyR)  row[$"{c} (Target only)"] = Val(right.Rows[i], c);
                        dt.Rows.Add(row); continue;
                    }
                    if (!hasR2)
                    {
                        row["Status"] = "⚠ Not in Target";
                        foreach (var c in common) row[$"{c} ({leftName})"] = Val(left.Rows[i], c);
                        foreach (var c in onlyL)  row[$"{c} (Source only)"] = Val(left.Rows[i], c);
                        dt.Rows.Add(row); continue;
                    }
                    var lr2 = left.Rows[i]; var rr2 = right.Rows[i];
                    bool allMatch = true;
                    foreach (var c in common)
                    {
                        string lv = Val(lr2, c), rv = Val(rr2, c);
                        bool   ok = string.Equals(lv.Trim(), rv.Trim(), StringComparison.OrdinalIgnoreCase);
                        row[$"{c} ({leftName})"]  = lv;
                        row[$"{c} ({rightName})"] = rv;
                        row[$"✓ {c}"]             = ok ? "✅" : "❌";
                        if (!ok) allMatch = false;
                    }
                    foreach (var c in onlyL) row[$"{c} (Source only)"] = Val(lr2, c);
                    foreach (var c in onlyR) row[$"{c} (Target only)"] = Val(rr2, c);
                    row["Status"] = allMatch ? "✅ Match" : "❌ Mismatch";
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        // ── Table Structure Compare ──────────────────────────────────────────────

        private async Task CompareStructureAsync()
        {
            var (lDb, lSch, lObj) = Left;
            var (rDb, rSch, rObj) = Right;

            using var cn = new SqlConnection(_connectionString);
            await cn.OpenAsync();
            var leftCols  = await FetchColumnsAsync(cn, lDb, lSch, lObj);
            var rightCols = await FetchColumnsAsync(cn, rDb, rSch, rObj);

            _result = BuildStructureCompare(leftCols, rightCols, lObj, rObj);
            ShowGrid();

            int same  = CountStatus("✅ Same");
            int diff  = CountStatus("❌ Different");
            int onlyL = CountStatus("⚠ Source only");
            int onlyR = CountStatus("⚠ Target only");
            lblSummary.Text =
                $"Structure: [{lDb}].{lObj}  vs  [{rDb}].{rObj}   " +
                $"✅ Same: {same}   ❌ Different: {diff}   ⚠ Source only: {onlyL}   ⚠ Target only: {onlyR}";
            lblSummary.ForeColor = diff + onlyL + onlyR > 0
                ? Color.FromArgb(255, 130, 130) : Color.FromArgb(100, 220, 100);
        }

        private static async Task<DataTable> FetchColumnsAsync(SqlConnection cn, string db, string sch, string tbl)
        {
            string s = sch.Replace("'", "''");
            string t = tbl.Replace("'", "''");
            string sql = $@"
SELECT COLUMN_NAME, DATA_TYPE,
  ISNULL(CAST(CHARACTER_MAXIMUM_LENGTH AS VARCHAR(10)),'') AS MAX_LEN,
  ISNULL(CAST(NUMERIC_PRECISION AS VARCHAR(10)),'')        AS NUM_P,
  ISNULL(CAST(NUMERIC_SCALE     AS VARCHAR(10)),'')        AS NUM_S,
  IS_NULLABLE,
  ISNULL(COLUMN_DEFAULT,'') AS COL_DEFAULT,
  ORDINAL_POSITION
FROM [{db}].INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA='{s}' AND TABLE_NAME='{t}'
ORDER BY ORDINAL_POSITION";
            var dt = new DataTable();
            using var da = new SqlDataAdapter(sql, cn);
            await Task.Run(() => da.Fill(dt));
            return dt;
        }

        private static DataTable BuildStructureCompare(DataTable left, DataTable right,
            string leftName, string rightName)
        {
            var dt = new DataTable();
            dt.Columns.Add("Status");
            dt.Columns.Add("Column Name");
            dt.Columns.Add($"DataType ({leftName})");
            dt.Columns.Add($"DataType ({rightName})");
            dt.Columns.Add("✓ Type");
            dt.Columns.Add($"Nullable ({leftName})");
            dt.Columns.Add($"Nullable ({rightName})");
            dt.Columns.Add("✓ Nullable");
            dt.Columns.Add($"Default ({leftName})");
            dt.Columns.Add($"Default ({rightName})");
            dt.Columns.Add("✓ Default");
            dt.Columns.Add($"Ordinal ({leftName})");
            dt.Columns.Add($"Ordinal ({rightName})");

            var lIdx = left.AsEnumerable()
                           .ToDictionary(r => r["COLUMN_NAME"].ToString()!, StringComparer.OrdinalIgnoreCase);
            var rIdx = right.AsEnumerable()
                            .ToDictionary(r => r["COLUMN_NAME"].ToString()!, StringComparer.OrdinalIgnoreCase);
            var all  = lIdx.Keys.Union(rIdx.Keys, StringComparer.OrdinalIgnoreCase)
                                .OrderBy(c => lIdx.TryGetValue(c, out var lr) ? (int)lr["ORDINAL_POSITION"]
                                            : rIdx.TryGetValue(c, out var rr) ? (int)rr["ORDINAL_POSITION"] + 10000
                                            : 99999);
            foreach (var col in all)
            {
                bool hasL = lIdx.TryGetValue(col, out var lc);
                bool hasR = rIdx.TryGetValue(col, out var rc);
                var row   = dt.NewRow();
                row["Column Name"] = col;

                if (!hasL) { row["Status"] = "⚠ Target only"; dt.Rows.Add(row); continue; }
                if (!hasR) { row["Status"] = "⚠ Source only"; dt.Rows.Add(row); continue; }

                string lType = FormatType(lc!), rType = FormatType(rc!);
                string lNull = lc!["IS_NULLABLE"].ToString()!, rNull = rc!["IS_NULLABLE"].ToString()!;
                string lDef  = lc["COL_DEFAULT"].ToString()!,  rDef  = rc["COL_DEFAULT"].ToString()!;
                string lOrd  = lc["ORDINAL_POSITION"].ToString()!, rOrd = rc["ORDINAL_POSITION"].ToString()!;

                bool typeOk = string.Equals(lType, rType, StringComparison.OrdinalIgnoreCase);
                bool nullOk = string.Equals(lNull, rNull, StringComparison.OrdinalIgnoreCase);
                bool defOk  = string.Equals(lDef,  rDef,  StringComparison.OrdinalIgnoreCase);

                row[$"DataType ({leftName})"]  = lType;  row[$"DataType ({rightName})"]  = rType;  row["✓ Type"]    = typeOk ? "✅" : "❌";
                row[$"Nullable ({leftName})"]  = lNull;  row[$"Nullable ({rightName})"]  = rNull;  row["✓ Nullable"] = nullOk ? "✅" : "❌";
                row[$"Default ({leftName})"]   = lDef;   row[$"Default ({rightName})"]   = rDef;   row["✓ Default"]  = defOk  ? "✅" : "❌";
                row[$"Ordinal ({leftName})"]   = lOrd;   row[$"Ordinal ({rightName})"]   = rOrd;
                row["Status"] = (typeOk && nullOk && defOk) ? "✅ Same" : "❌ Different";
                dt.Rows.Add(row);
            }
            return dt;
        }

        private static string FormatType(DataRow r)
        {
            string t  = r["DATA_TYPE"].ToString()!;
            string ml = r["MAX_LEN"].ToString()!;
            string np = r["NUM_P"].ToString()!;
            string ns = r["NUM_S"].ToString()!;
            if (!string.IsNullOrEmpty(ml) && ml != "-1") return $"{t}({ml})";
            if (ml == "-1") return $"{t}(MAX)";
            if (!string.IsNullOrEmpty(np) && !string.IsNullOrEmpty(ns)) return $"{t}({np},{ns})";
            if (!string.IsNullOrEmpty(np)) return $"{t}({np})";
            return t;
        }

        // ── Stored Procedure Compare ─────────────────────────────────────────────

        private async Task CompareProcAsync()
        {
            var (lDb, lSch, lObj) = Left;
            var (rDb, rSch, rObj) = Right;

            using var cn = new SqlConnection(_connectionString);
            await cn.OpenAsync();
            string? leftDef  = await FetchProcDefAsync(cn, lDb, lSch, lObj);
            string? rightDef = await FetchProcDefAsync(cn, rDb, rSch, rObj);

            if (leftDef == null && rightDef == null)
            {
                lblSummary.Text = "Neither procedure was found.";
                return;
            }

            ShowDiff(leftDef ?? "(procedure not found)", rightDef ?? "(procedure not found)", lObj, rObj, lDb, rDb);
        }

        private static async Task<string?> FetchProcDefAsync(SqlConnection cn, string db, string sch, string proc)
        {
            using var cmd = new SqlCommand(
                $"SELECT OBJECT_DEFINITION(OBJECT_ID('[{db}].[{sch}].[{proc}]'))", cn);
            var v = await cmd.ExecuteScalarAsync();
            return v == DBNull.Value || v == null ? null : v.ToString();
        }

        private void ShowDiff(string left, string right,
            string leftName, string rightName, string leftDb, string rightDb)
        {
            var lLines = left.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
            var rLines = right.Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
            var ops    = ComputeDiff(lLines, rLines);

            txtDiff.Clear();
            // Header
            AppendDiff($"--- Source: [{leftDb}].{leftName}\n",  Color.FromArgb(150, 150, 210));
            AppendDiff($"+++ Target: [{rightDb}].{rightName}\n", Color.FromArgb(150, 210, 150));
            AppendDiff("\n", Color.FromArgb(212, 212, 212));

            foreach (var (op, line) in ops)
            {
                Color bg = op switch
                {
                    '+' => Color.FromArgb(20, 60, 20),
                    '-' => Color.FromArgb(70, 20, 20),
                    _   => Color.FromArgb(20, 20, 20)
                };
                Color fg = op switch
                {
                    '+' => Color.FromArgb(140, 230, 140),
                    '-' => Color.FromArgb(230, 130, 130),
                    _   => Color.FromArgb(212, 212, 212)
                };
                string prefix = op switch { '+' => "+ ", '-' => "- ", _ => "  " };
                AppendDiff(prefix + line + "\n", fg, bg);
            }

            int adds = ops.Count(o => o.op == '+');
            int dels = ops.Count(o => o.op == '-');

            txtDiff.Visible     = true;
            grid.Visible        = false;
            btnExportCsv.Enabled  = false;
            btnExportXlsx.Enabled = false;

            bool identical = adds == 0 && dels == 0;
            lblSummary.Text = identical
                ? $"Stored Procedure: {leftName}  vs  {rightName}   ✅ Identical"
                : $"Stored Procedure: {leftName}  vs  {rightName}   ❌ +{adds} lines added  -{dels} lines removed";
            lblSummary.ForeColor = identical
                ? Color.FromArgb(100, 220, 100) : Color.FromArgb(255, 130, 130);
        }

        // Myers-style LCS diff — fallback to simple for very large inputs
        private static List<(char op, string line)> ComputeDiff(string[] a, string[] b)
        {
            if (a.Length > 800 || b.Length > 800) return SimpleDiff(a, b);

            int m = a.Length, n = b.Length;
            int[,] lcs = new int[m + 1, n + 1];
            for (int i = m - 1; i >= 0; i--)
                for (int j = n - 1; j >= 0; j--)
                    lcs[i, j] = string.Equals(a[i], b[j], StringComparison.Ordinal)
                        ? lcs[i + 1, j + 1] + 1
                        : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);

            var result = new List<(char, string)>();
            int ai = 0, bi = 0;
            while (ai < m || bi < n)
            {
                if (ai < m && bi < n && string.Equals(a[ai], b[bi], StringComparison.Ordinal))
                    { result.Add((' ', a[ai])); ai++; bi++; }
                else if (bi < n && (ai >= m || lcs[ai, bi + 1] >= lcs[ai + 1, bi]))
                    { result.Add(('+', b[bi])); bi++; }
                else
                    { result.Add(('-', a[ai])); ai++; }
            }
            return result;
        }

        private static List<(char op, string line)> SimpleDiff(string[] a, string[] b)
        {
            var r = new List<(char, string)>();
            foreach (var l in a) r.Add(('-', l));
            foreach (var l in b) r.Add(('+', l));
            return r;
        }

        private void AppendDiff(string text, Color fg, Color? bg = null)
        {
            int start = txtDiff.TextLength;
            txtDiff.AppendText(text);
            txtDiff.Select(start, text.Length);
            if (bg.HasValue) txtDiff.SelectionBackColor = bg.Value;
            txtDiff.SelectionColor  = fg;
            txtDiff.SelectionLength = 0;
        }

        // ── Grid ─────────────────────────────────────────────────────────────────

        private void ShowGrid()
        {
            grid.DataSource   = _result;
            grid.Visible      = true;
            txtDiff.Visible   = false;
            btnExportCsv.Enabled  = true;
            btnExportXlsx.Enabled = true;
            // Widen Status column, narrow check columns
            if (grid.Columns.Contains("Status")) grid.Columns["Status"]!.Width = 140;
            foreach (DataGridViewColumn col in grid.Columns)
                if (col.Name.StartsWith("✓ ")) col.Width = 38;
        }

        private void Grid_RowPrePaint(object? sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= grid.Rows.Count) return;
            if (grid.Rows[e.RowIndex].DataBoundItem is not DataRowView drv) return;
            string status = drv.Row["Status"]?.ToString() ?? "";
            grid.Rows[e.RowIndex].DefaultCellStyle.BackColor = status switch
            {
                "✅ Match" or "✅ Same"           => Color.FromArgb(18, 50, 22),
                "❌ Mismatch" or "❌ Different"   => Color.FromArgb(62, 18, 18),
                _                                 => Color.FromArgb(55, 44, 18)
            };
        }

        // ── Export ───────────────────────────────────────────────────────────────

        private void ExportCsv()
        {
            if (_result == null || _result.Rows.Count == 0) return;
            using var dlg = new SaveFileDialog { Filter = "CSV|*.csv", FileName = "db_compare.csv" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", _result.Columns.Cast<DataColumn>().Select(c => Esc(c.ColumnName))));
            foreach (DataRow row in _result.Rows)
                sb.AppendLine(string.Join(",", row.ItemArray.Select(v => Esc(v?.ToString() ?? ""))));
            File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
            MessageBox.Show("Exported!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportExcel()
        {
            if (_result == null || _result.Rows.Count == 0) return;
            using var dlg = new SaveFileDialog { Filter = "Excel|*.xlsx", FileName = "db_compare.xlsx" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Compare");
            for (int c = 0; c < _result.Columns.Count; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = _result.Columns[c].ColumnName;
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor  = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(37, 37, 38);
            }
            for (int r = 0; r < _result.Rows.Count; r++)
            {
                string status = _result.Rows[r]["Status"]?.ToString() ?? "";
                var bg = status switch
                {
                    "✅ Match" or "✅ Same"         => XLColor.FromArgb(198, 239, 206),
                    "❌ Mismatch" or "❌ Different" => XLColor.FromArgb(255, 199, 206),
                    _                               => XLColor.FromArgb(255, 235, 156)
                };
                for (int c = 0; c < _result.Columns.Count; c++)
                {
                    var cell = ws.Cell(r + 2, c + 1);
                    cell.Value = _result.Rows[r][c]?.ToString() ?? "";
                    cell.Style.Fill.BackgroundColor = bg;
                    if (_result.Columns[c].ColumnName.StartsWith("✓ ") && cell.Value.ToString() == "❌")
                        cell.Style.Font.FontColor = XLColor.Red;
                }
            }
            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
            wb.SaveAs(dlg.FileName);
            MessageBox.Show("Exported!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── Misc Helpers ─────────────────────────────────────────────────────────

        private int CountStatus(string status) =>
            _result?.AsEnumerable().Count(r => r["Status"].ToString() == status) ?? 0;

        private static string Val(DataRow row, string col)
        {
            try { return row[col]?.ToString() ?? ""; }
            catch { return ""; }
        }

        private static Dictionary<string, DataRow> BuildLookup(DataTable dt, string keyCol) =>
            dt.AsEnumerable()
              .GroupBy(r => Val(r, keyCol), StringComparer.OrdinalIgnoreCase)
              .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        private static RadioButton MakeRdo(string text, int x, int y) => new RadioButton
        {
            Text = text, Location = new Point(x, y), AutoSize = true,
            ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9.5f)
        };

        private static GroupBox MakeGroup(string title, int x, int y) => new GroupBox
        {
            Text = title, Location = new Point(x, y),
            ForeColor = Color.FromArgb(150, 150, 150), Font = new Font("Segoe UI", 8.5f)
        };

        private static Button MakeBtn(string text, int x, int y, int w)
        {
            var b = new Button
            {
                Text = text, Location = new Point(x, y), Width = w, Height = 27,
                FlatStyle = FlatStyle.Flat, BackColor = Color.FromArgb(65, 65, 68),
                ForeColor = Color.White, Font = new Font("Segoe UI", 9f)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private static ComboBox MakeCombo(int x, int y, int w) => new ComboBox
        {
            Location = new Point(x, y), Width = w, DropDownStyle = ComboBoxStyle.DropDownList,
            BackColor = Color.FromArgb(55, 55, 55), ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f)
        };

        private static Label MakeLabel(string text, int x, int y) => new Label
        {
            Text = text, Location = new Point(x, y), AutoSize = true,
            ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9.5f)
        };

        private static Label MakeSmLabel(string text, int x, int y, int w) => new Label
        {
            Text = text, Location = new Point(x, y), Width = w,
            ForeColor = Color.FromArgb(140, 140, 140), Font = new Font("Segoe UI", 8.5f)
        };

        private static string Esc(string s) =>
            s.Contains(',') || s.Contains('"') || s.Contains('\n')
                ? $"\"{s.Replace("\"", "\"\"")}\"" : s;
    }
}
