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
    public class DeploymentCheckerDialog : Form
    {
        private readonly string _connectionString;

        // Top bar
        private ComboBox    cboDatabase  = null!;
        private DateTimePicker dtpFrom   = null!;
        private DateTimePicker dtpTo     = null!;
        private Button      btnScan      = null!;

        // Type filters
        private CheckBox chkTables    = null!;
        private CheckBox chkProcs     = null!;
        private CheckBox chkViews     = null!;
        private CheckBox chkFunctions = null!;
        private CheckBox chkTriggers  = null!;

        // Object list
        private ListView    lvObjects   = null!;
        private Label       lblCount    = null!;
        private Button      btnAll      = null!;
        private Button      btnNone     = null!;

        // Script panel
        private RichTextBox txtScript   = null!;
        private Button      btnGenerate = null!;
        private Button      btnCopy     = null!;
        private Button      btnSaveSql  = null!;
        private Button      btnExportList = null!;

        private sealed record DbObject(
            string Type, string TypeDesc, string Schema, string Name,
            DateTime CreateDate, DateTime ModifyDate, string Status);

        private List<DbObject> _objects = new();

        public DeploymentCheckerDialog(string connectionString)
        {
            _connectionString = connectionString;
            BuildUI();
        }

        // ── UI ──────────────────────────────────────────────────────────────────

        private void BuildUI()
        {
            Text          = "Deployment Checker — Changed Objects";
            Size          = new Size(1280, 820);
            MinimumSize   = new Size(1000, 640);
            StartPosition = FormStartPosition.CenterParent;
            BackColor     = Color.FromArgb(30, 30, 30);

            // ── Top scan bar ─────────────────────────────────────────────────────
            var pnlTop = new Panel
            {
                Dock = DockStyle.Top, Height = 44,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            pnlTop.Controls.Add(MakeLabel("Database:", 8, 14));
            cboDatabase = MakeCombo(80, 10, 180);
            pnlTop.Controls.Add(cboDatabase);

            pnlTop.Controls.Add(MakeLabel("From:", 272, 14));
            dtpFrom = new DateTimePicker
            {
                Location = new Point(310, 9), Width = 150, Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddDays(-30),
                CalendarForeColor = Color.Black, Font = new Font("Segoe UI", 9f)
            };
            pnlTop.Controls.Add(dtpFrom);

            pnlTop.Controls.Add(MakeLabel("To:", 470, 14));
            dtpTo = new DateTimePicker
            {
                Location = new Point(490, 9), Width = 150, Format = DateTimePickerFormat.Short,
                Value = DateTime.Today,
                Font = new Font("Segoe UI", 9f)
            };
            pnlTop.Controls.Add(dtpTo);

            btnScan = MakeBtn("🔍  Scan", 654, 8, 110);
            btnScan.BackColor = Color.FromArgb(0, 122, 204);
            btnScan.Click += BtnScan_Click;
            pnlTop.Controls.Add(btnScan);

            // ── Filter bar ───────────────────────────────────────────────────────
            var pnlFilter = new Panel
            {
                Dock = DockStyle.Top, Height = 34,
                BackColor = Color.FromArgb(37, 37, 38)
            };
            pnlFilter.Controls.Add(MakeLabel("Show:", 8, 9));
            chkTables    = MakeChk("Tables",            65,  8, true);
            chkProcs     = MakeChk("Stored Procs",     155,  8, true);
            chkViews     = MakeChk("Views",            270,  8, true);
            chkFunctions = MakeChk("Functions",        340,  8, true);
            chkTriggers  = MakeChk("Triggers",         430,  8, true);
            pnlFilter.Controls.AddRange(new Control[]
                { chkTables, chkProcs, chkViews, chkFunctions, chkTriggers });

            // ── Split: list (left) / script (right) ──────────────────────────────
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill, Orientation = Orientation.Vertical,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            BuildListPanel(split.Panel1);
            BuildScriptPanel(split.Panel2);

            SuspendLayout();
            Controls.Add(split);
            Controls.Add(pnlFilter);
            Controls.Add(pnlTop);
            ResumeLayout();

            Load += async (s, e) =>
            {
                split.Panel1MinSize    = 380;
                split.Panel2MinSize    = 340;
                split.SplitterDistance = Math.Max(380, (int)(split.Width * 0.48));
                await LoadDatabasesAsync();
            };
        }

        private void BuildListPanel(SplitterPanel pnl)
        {
            pnl.BackColor = Color.FromArgb(30, 30, 30);

            // Header
            var pnlHdr = new Panel
            {
                Dock = DockStyle.Top, Height = 32,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            lblCount = MakeLabel("Scan to find changed objects.", 8, 9);
            btnAll   = MakeBtn("Select All",  0, 4, 90);
            btnNone  = MakeBtn("Select None", 0, 4, 90);
            btnAll.Enabled = btnNone.Enabled = false;
            btnAll.Click  += (s, e) => SetAllChecked(true);
            btnNone.Click += (s, e) => SetAllChecked(false);

            pnlHdr.Controls.AddRange(new Control[] { lblCount, btnAll, btnNone });
            pnlHdr.Resize += (s, e) =>
            {
                btnNone.Location = new Point(pnlHdr.Width - 98, 4);
                btnAll.Location  = new Point(pnlHdr.Width - 194, 4);
            };

            // ListView
            lvObjects = new ListView
            {
                Dock = DockStyle.Fill, View = View.Details,
                CheckBoxes = true, FullRowSelect = true, MultiSelect = true,
                BackColor = Color.FromArgb(30, 30, 30), ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None, Font = new Font("Consolas", 9f),
                OwnerDraw = true, ShowItemToolTips = true
            };
            lvObjects.Columns.Add("Status",   62);
            lvObjects.Columns.Add("Type",     48);
            lvObjects.Columns.Add("Schema",   64);
            lvObjects.Columns.Add("Name",    230);
            lvObjects.Columns.Add("Modified", 130);
            lvObjects.Columns.Add("Created",  130);

            lvObjects.DrawColumnHeader += LvDrawColumnHeader;
            lvObjects.DrawItem         += (s, e) => { };
            lvObjects.DrawSubItem      += LvDrawSubItem;
            lvObjects.ItemChecked      += (s, e) => UpdateGenerateBtn();
            lvObjects.ColumnClick      += LvColumnClick;

            // Bottom bar
            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom, Height = 38,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            btnExportList = MakeBtn("Export List (Excel)", 6, 6, 148);
            btnExportList.BackColor = Color.FromArgb(20, 100, 45);
            btnExportList.Enabled   = false;
            btnExportList.Click    += (s, e) => ExportList();
            pnlBottom.Controls.Add(btnExportList);

            pnl.Controls.Add(lvObjects);
            pnl.Controls.Add(pnlBottom);
            pnl.Controls.Add(pnlHdr);
        }

        private void BuildScriptPanel(SplitterPanel pnl)
        {
            pnl.BackColor = Color.FromArgb(30, 30, 30);

            var pnlHdr = new Panel
            {
                Dock = DockStyle.Top, Height = 32,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            btnGenerate = MakeBtn("⚙  Generate Script for Checked", 6, 4, 220);
            btnGenerate.BackColor = Color.FromArgb(40, 70, 100);
            btnGenerate.Enabled   = false;
            btnGenerate.Click    += BtnGenerate_Click;
            pnlHdr.Controls.Add(btnGenerate);

            txtScript = new RichTextBox
            {
                Dock = DockStyle.Fill, ReadOnly = true,
                Font = new Font("Consolas", 10f),
                BackColor = Color.FromArgb(20, 20, 20), ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None, ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false
            };

            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom, Height = 38,
                BackColor = Color.FromArgb(45, 45, 48)
            };
            btnCopy    = MakeBtn("📋 Copy",        6,   6, 90);
            btnSaveSql = MakeBtn("💾 Save .sql", 102,   6, 110);
            btnCopy.Enabled = btnSaveSql.Enabled = false;
            btnCopy.Click    += (s, e) => { if (txtScript.TextLength > 0) Clipboard.SetText(txtScript.Text); };
            btnSaveSql.Click += BtnSaveSql_Click;
            pnlBottom.Controls.AddRange(new Control[] { btnCopy, btnSaveSql });

            pnl.Controls.Add(txtScript);
            pnl.Controls.Add(pnlBottom);
            pnl.Controls.Add(pnlHdr);
        }

        // ── Database Loading ─────────────────────────────────────────────────────

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
                string cur = cn.Database;
                cboDatabase.Items.Clear();
                foreach (DataRow r in dt.Rows) cboDatabase.Items.Add(r[0].ToString()!);
                int idx = cboDatabase.Items.IndexOf(cur);
                cboDatabase.SelectedIndex = idx >= 0 ? idx : cboDatabase.Items.Count > 0 ? 0 : -1;
            }
            catch (Exception ex) { lblCount.Text = $"Error: {ex.Message}"; }
        }

        // ── Scan ─────────────────────────────────────────────────────────────────

        private async void BtnScan_Click(object? sender, EventArgs e)
        {
            if (cboDatabase.SelectedItem == null) return;
            btnScan.Enabled = false;
            lblCount.Text   = "Scanning…";
            lvObjects.Items.Clear();
            txtScript.Clear();
            _objects.Clear();

            try
            {
                string db = cboDatabase.SelectedItem.ToString()!;
                _objects = await ScanObjectsAsync(db, dtpFrom.Value.Date, dtpTo.Value.Date.AddDays(1).AddSeconds(-1));
                PopulateList();
            }
            catch (Exception ex)
            {
                lblCount.Text = $"Error: {ex.Message}";
            }
            finally { btnScan.Enabled = true; }
        }

        private async Task<List<DbObject>> ScanObjectsAsync(string db, DateTime from, DateTime to)
        {
            var types = new List<string>();
            if (chkTables.Checked)    types.Add("'U'");
            if (chkProcs.Checked)     types.Add("'P'");
            if (chkViews.Checked)     types.Add("'V'");
            if (chkFunctions.Checked) types.AddRange(new[] { "'FN'","'IF'","'TF'" });
            if (chkTriggers.Checked)  types.Add("'TR'");

            if (types.Count == 0) return new List<DbObject>();

            string sql = $@"
USE [{db}];
SELECT
    o.type,
    RTRIM(o.type_desc) AS type_desc,
    SCHEMA_NAME(o.schema_id) AS schema_name,
    o.name,
    o.create_date,
    o.modify_date,
    CASE WHEN o.create_date >= @from THEN 'New' ELSE 'Modified' END AS status
FROM sys.objects o
WHERE o.is_ms_shipped = 0
  AND o.type IN ({string.Join(",", types)})
  AND o.modify_date >= @from
  AND o.modify_date <= @to
ORDER BY o.modify_date DESC";

            using var cn = new SqlConnection(_connectionString);
            await cn.OpenAsync();
            using var cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("@from", from);
            cmd.Parameters.AddWithValue("@to",   to);

            var list = new List<DbObject>();
            using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                list.Add(new DbObject(
                    Type:       rd.GetString(0).Trim(),
                    TypeDesc:   rd.GetString(1),
                    Schema:     rd.GetString(2),
                    Name:       rd.GetString(3),
                    CreateDate: rd.GetDateTime(4),
                    ModifyDate: rd.GetDateTime(5),
                    Status:     rd.GetString(6)
                ));
            }
            return list;
        }

        private void PopulateList()
        {
            lvObjects.Items.Clear();
            foreach (var o in _objects)
            {
                string typeShort = o.Type switch
                {
                    "U"  => "TABLE",
                    "P"  => "PROC",
                    "V"  => "VIEW",
                    "FN" or "IF" or "TF" => "FUNC",
                    "TR" => "TRIGGER",
                    _    => o.Type
                };
                var item = new ListViewItem(o.Status == "New" ? "🆕 New" : "✏️ Modified")
                {
                    Checked = true, Tag = o,
                    ForeColor = o.Status == "New"
                        ? Color.FromArgb(100, 220, 100)
                        : Color.FromArgb(100, 180, 255)
                };
                item.SubItems.Add(typeShort);
                item.SubItems.Add(o.Schema);
                item.SubItems.Add(o.Name);
                item.SubItems.Add(o.ModifyDate.ToString("yyyy-MM-dd HH:mm:ss"));
                item.SubItems.Add(o.CreateDate.ToString("yyyy-MM-dd HH:mm:ss"));
                item.ToolTipText = $"{o.Schema}.{o.Name}\nType: {o.TypeDesc}\nModified: {o.ModifyDate:yyyy-MM-dd HH:mm:ss}";
                lvObjects.Items.Add(item);
            }

            int count = _objects.Count;
            lblCount.Text = count == 0
                ? "No changes found in this date range."
                : $"{count} object(s) changed  —  all checked";

            bool any = count > 0;
            btnAll.Enabled = btnNone.Enabled = btnExportList.Enabled = any;
            UpdateGenerateBtn();
        }

        private void SetAllChecked(bool v)
        {
            foreach (ListViewItem item in lvObjects.Items) item.Checked = v;
        }

        private void UpdateGenerateBtn()
        {
            int n = lvObjects.CheckedItems.Count;
            btnGenerate.Enabled = n > 0;
            btnGenerate.Text    = n > 0
                ? $"⚙  Generate Script ({n} objects)"
                : "⚙  Generate Script for Checked";
        }

        // ── Script Generation ────────────────────────────────────────────────────

        private async void BtnGenerate_Click(object? sender, EventArgs e)
        {
            if (cboDatabase.SelectedItem == null) return;
            btnGenerate.Enabled = false;
            txtScript.Clear();
            ScriptAppend($"-- ============================================================\n");
            ScriptAppend($"-- Deployment Script\n");
            ScriptAppend($"-- Database : {cboDatabase.SelectedItem}\n");
            ScriptAppend($"-- Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");
            ScriptAppend($"-- Objects  : {lvObjects.CheckedItems.Count}\n");
            ScriptAppend($"-- ============================================================\n\n");
            ScriptAppend($"USE [{cboDatabase.SelectedItem}];\nGO\n\n");

            try
            {
                using var cn = new SqlConnection(_connectionString);
                await cn.OpenAsync();

                foreach (ListViewItem item in lvObjects.CheckedItems)
                {
                    if (item.Tag is not DbObject o) continue;
                    string script = o.Type switch
                    {
                        "U"                    => await GenTableScriptAsync(cn, o),
                        "P" or "V" or "FN" or "IF" or "TF" => await GenModuleScriptAsync(cn, o),
                        "TR"                   => await GenModuleScriptAsync(cn, o),
                        _                      => $"-- (Script generation not supported for type {o.TypeDesc})\n"
                    };
                    ScriptAppend(script + "\n");
                }

                btnCopy.Enabled    = true;
                btnSaveSql.Enabled = true;
            }
            catch (Exception ex)
            {
                ScriptAppend($"\n-- ERROR: {ex.Message}\n");
            }
            finally { btnGenerate.Enabled = true; }
        }

        // Stored Procedure / View / Function / Trigger
        private static async Task<string> GenModuleScriptAsync(SqlConnection cn, DbObject o)
        {
            using var cmd = new SqlCommand(
                $"SELECT OBJECT_DEFINITION(OBJECT_ID('[{o.Schema}].[{o.Name}]'))", cn);
            var def = (await cmd.ExecuteScalarAsync())?.ToString();
            if (string.IsNullOrEmpty(def))
                return $"-- WARNING: Definition not found for [{o.Schema}].[{o.Name}]\n";

            // Convert CREATE to CREATE OR ALTER for procs/functions, wrap others in IF EXISTS DROP
            string header =
                $"-- ------------------------------------------------------------\n" +
                $"-- {o.TypeDesc}: [{o.Schema}].[{o.Name}]\n" +
                $"-- Modified : {o.ModifyDate:yyyy-MM-dd HH:mm:ss}\n" +
                $"-- Status   : {o.Status}\n" +
                $"-- ------------------------------------------------------------\n";

            // Try CREATE OR ALTER (SQL Server 2016+) for procs & scalar functions
            if (o.Type is "P" or "FN" or "IF" or "TF")
            {
                string altered = System.Text.RegularExpressions.Regex.Replace(
                    def,
                    @"\bCREATE\s+(PROCEDURE|PROC|FUNCTION)\b",
                    "CREATE OR ALTER $1",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase,
                    TimeSpan.FromSeconds(2));
                return header + altered + "\nGO\n";
            }

            // Views and triggers: DROP + CREATE pattern
            string dropType = o.Type switch
            {
                "V"  => "VIEW",
                "TR" => "TRIGGER",
                _    => "OBJECT"
            };
            string dropSql =
                $"IF OBJECT_ID(N'[{o.Schema}].[{o.Name}]') IS NOT NULL\n" +
                $"    DROP {dropType} [{o.Schema}].[{o.Name}];\nGO\n";
            return header + dropSql + def + "\nGO\n";
        }

        // Table: generate CREATE TABLE + PK
        private static async Task<string> GenTableScriptAsync(SqlConnection cn, DbObject o)
        {
            // Columns
            string colSql = $@"
SELECT c.COLUMN_NAME, c.DATA_TYPE,
       c.CHARACTER_MAXIMUM_LENGTH, c.NUMERIC_PRECISION, c.NUMERIC_SCALE,
       c.IS_NULLABLE, c.COLUMN_DEFAULT,
       COLUMNPROPERTY(OBJECT_ID('{o.Schema}.{o.Name}'), c.COLUMN_NAME, 'IsIdentity') AS IS_IDENTITY,
       IDENT_SEED('{o.Schema}.{o.Name}')      AS SEED,
       IDENT_INCR('{o.Schema}.{o.Name}')      AS INCR
FROM INFORMATION_SCHEMA.COLUMNS c
WHERE c.TABLE_SCHEMA = '{o.Schema.Replace("'","''")}' AND c.TABLE_NAME = '{o.Name.Replace("'","''")}'
ORDER BY c.ORDINAL_POSITION";
            var colDt = new DataTable();
            using var colCmd = new SqlCommand(colSql, cn);
            using (var rd = await colCmd.ExecuteReaderAsync())
                colDt.Load(rd);

            // PKs
            string pkSql = $@"
SELECT kc.COLUMN_NAME
FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE kc
JOIN INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
    ON kc.CONSTRAINT_NAME = tc.CONSTRAINT_NAME
   AND kc.TABLE_SCHEMA = tc.TABLE_SCHEMA
WHERE tc.TABLE_SCHEMA = '{o.Schema.Replace("'","''")}' AND tc.TABLE_NAME = '{o.Name.Replace("'","''")}'
  AND tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
ORDER BY kc.ORDINAL_POSITION";
            var pkCols = new List<string>();
            using var pkCmd = new SqlCommand(pkSql, cn);
            using (var rd = await pkCmd.ExecuteReaderAsync())
                while (await rd.ReadAsync()) pkCols.Add(rd.GetString(0));

            var sb = new StringBuilder();
            sb.AppendLine("-- ------------------------------------------------------------");
            sb.AppendLine($"-- TABLE: [{o.Schema}].[{o.Name}]");
            sb.AppendLine($"-- Modified: {o.ModifyDate:yyyy-MM-dd HH:mm:ss}  Status: {o.Status}");
            sb.AppendLine("-- NOTE: This is the current full definition for reference.");
            sb.AppendLine("--       For ALTER scripts, compare with target using Compare DB.");
            sb.AppendLine("-- ------------------------------------------------------------");
            sb.AppendLine($"IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES");
            sb.AppendLine($"              WHERE TABLE_SCHEMA='{o.Schema}' AND TABLE_NAME='{o.Name}')");
            sb.AppendLine("BEGIN");
            sb.AppendLine($"    CREATE TABLE [{o.Schema}].[{o.Name}] (");

            var colDefs = new List<string>();
            foreach (DataRow r in colDt.Rows)
            {
                string colName = $"[{r["COLUMN_NAME"]}]";
                string typeDef = FormatColType(r);
                bool identity  = r["IS_IDENTITY"] is int id && id == 1;
                string identStr = identity
                    ? $" IDENTITY({r["SEED"]},{r["INCR"]})"
                    : "";
                bool nullable  = r["IS_NULLABLE"].ToString() == "YES";
                string nullStr = nullable ? " NULL" : " NOT NULL";
                string defStr  = "";
                string? defVal = r["COLUMN_DEFAULT"]?.ToString();
                if (!string.IsNullOrEmpty(defVal)) defStr = $" DEFAULT {defVal}";
                colDefs.Add($"        {colName} {typeDef}{identStr}{nullStr}{defStr}");
            }

            if (pkCols.Count > 0)
                colDefs.Add($"        CONSTRAINT [PK_{o.Name}] PRIMARY KEY ({string.Join(", ", pkCols.Select(c => $"[{c}]"))})");

            sb.AppendLine(string.Join(",\n", colDefs));
            sb.AppendLine("    );");
            sb.AppendLine("END");
            sb.AppendLine("GO");
            return sb.ToString();
        }

        private static string FormatColType(DataRow r)
        {
            string t  = r["DATA_TYPE"].ToString()!;
            int?   ml = r["CHARACTER_MAXIMUM_LENGTH"] is DBNull ? null : Convert.ToInt32(r["CHARACTER_MAXIMUM_LENGTH"]);
            int?   np = r["NUMERIC_PRECISION"]        is DBNull ? null : Convert.ToInt32(r["NUMERIC_PRECISION"]);
            int?   ns = r["NUMERIC_SCALE"]            is DBNull ? null : Convert.ToInt32(r["NUMERIC_SCALE"]);
            if (ml.HasValue) return ml == -1 ? $"{t}(MAX)" : $"{t}({ml})";
            if (np.HasValue && ns.HasValue && t is "decimal" or "numeric") return $"{t}({np},{ns})";
            return t;
        }

        // ── ListView Draw ────────────────────────────────────────────────────────

        private int _sortCol = 4; // Modified date
        private bool _sortAsc = false;

        private void LvColumnClick(object? sender, ColumnClickEventArgs e)
        {
            if (e.Column == _sortCol) _sortAsc = !_sortAsc;
            else { _sortCol = e.Column; _sortAsc = true; }
            _objects = (_sortAsc
                ? _objects.OrderBy(SortKey(_sortCol))
                : _objects.OrderByDescending(SortKey(_sortCol)))
                .ToList();
            PopulateList();
        }

        private static Func<DbObject, object> SortKey(int col) => col switch
        {
            1 => o => o.Type,
            2 => o => o.Schema,
            3 => o => o.Name,
            4 => o => o.ModifyDate,
            5 => o => o.CreateDate,
            _ => o => o.Status
        };

        private void LvDrawColumnHeader(object? sender, DrawListViewColumnHeaderEventArgs e)
        {
            using var bg = new SolidBrush(Color.FromArgb(45, 45, 48));
            e.Graphics.FillRectangle(bg, e.Bounds);
            string title = e.Header!.Text + (e.ColumnIndex == _sortCol ? (_sortAsc ? " ▲" : " ▼") : "");
            TextRenderer.DrawText(e.Graphics, title, lvObjects.Font, e.Bounds,
                Color.FromArgb(160, 160, 160),
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
        }

        private void LvDrawSubItem(object? sender, DrawListViewSubItemEventArgs e)
        {
            if (e.Item == null) return;
            bool sel = e.Item.Selected;
            var bg = sel
                ? Color.FromArgb(0, 90, 160)
                : e.ItemIndex % 2 == 0 ? Color.FromArgb(30, 30, 30) : Color.FromArgb(34, 34, 36);
            var fg = sel ? Color.White : e.Item.ForeColor;

            using var bgBrush = new SolidBrush(bg);
            e.Graphics.FillRectangle(bgBrush, e.Bounds);

            var rect = new Rectangle(e.Bounds.X + 3, e.Bounds.Y, e.Bounds.Width - 3, e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, e.SubItem?.Text ?? "", lvObjects.Font, rect, fg,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        // ── Script View ──────────────────────────────────────────────────────────

        private void ScriptAppend(string text)
        {
            // Color SQL keywords
            bool isComment = text.TrimStart().StartsWith("--");
            Color fg = isComment
                ? Color.FromArgb(100, 150, 100)
                : Color.FromArgb(212, 212, 212);

            int start = txtScript.TextLength;
            txtScript.AppendText(text);
            txtScript.Select(start, text.Length);
            txtScript.SelectionColor  = fg;
            txtScript.SelectionLength = 0;
            txtScript.ScrollToCaret();
        }

        // ── Export & Save ────────────────────────────────────────────────────────

        private void ExportList()
        {
            if (_objects.Count == 0) return;
            using var dlg = new SaveFileDialog
            {
                Filter   = "Excel|*.xlsx",
                FileName = $"deploy_check_{cboDatabase.SelectedItem}_{DateTime.Now:yyyyMMdd}.xlsx"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Changed Objects");

            string[] headers = { "Status", "Type", "Schema", "Name", "Modified", "Created", "Checked" };
            for (int c = 0; c < headers.Length; c++)
            {
                var cell = ws.Cell(1, c + 1);
                cell.Value = headers[c];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromArgb(37, 37, 38);
            }

            int row = 2;
            foreach (ListViewItem item in lvObjects.Items)
            {
                if (item.Tag is not DbObject o) continue;
                ws.Cell(row, 1).Value = o.Status;
                ws.Cell(row, 2).Value = item.SubItems[1].Text;
                ws.Cell(row, 3).Value = o.Schema;
                ws.Cell(row, 4).Value = o.Name;
                ws.Cell(row, 5).Value = o.ModifyDate.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cell(row, 6).Value = o.CreateDate.ToString("yyyy-MM-dd HH:mm:ss");
                ws.Cell(row, 7).Value = item.Checked ? "✓" : "";

                var bg = o.Status == "New"
                    ? XLColor.FromArgb(198, 239, 206)
                    : XLColor.FromArgb(220, 230, 255);
                for (int c = 1; c <= 7; c++)
                    ws.Cell(row, c).Style.Fill.BackgroundColor = bg;
                row++;
            }

            ws.Columns().AdjustToContents();
            ws.SheetView.FreezeRows(1);
            wb.SaveAs(dlg.FileName);
            MessageBox.Show("Checklist exported!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSaveSql_Click(object? sender, EventArgs e)
        {
            if (txtScript.TextLength == 0) return;
            using var dlg = new SaveFileDialog
            {
                Filter   = "SQL Script|*.sql",
                FileName = $"deploy_{cboDatabase.SelectedItem}_{DateTime.Now:yyyyMMdd}.sql"
            };
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            File.WriteAllText(dlg.FileName, txtScript.Text, Encoding.UTF8);
            MessageBox.Show("Script saved!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ── UI Helpers ───────────────────────────────────────────────────────────

        private static Button MakeBtn(string text, int x, int y, int w)
        {
            var b = new Button
            {
                Text = text, Location = new Point(x, y), Width = w, Height = 26,
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
            ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9f)
        };

        private static CheckBox MakeChk(string text, int x, int y, bool chk) => new CheckBox
        {
            Text = text, Location = new Point(x, y), AutoSize = true, Checked = chk,
            ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 9f)
        };
    }
}
