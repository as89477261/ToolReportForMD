using System.Drawing;
using System.Windows.Forms;

namespace SqlQueryTool
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // Top toolbar controls
        private Panel pnlToolbar;
        private Panel picStatus;
        private Label lblConnInfo;
        private Button btnConnect;
        private Button btnDisconnect;

        // Query area
        private RichTextBox rtbQuery;
        private Panel pnlQueryToolbar;
        private Button btnExecute;
        private Button btnExecuteNonQuery;
        private Button btnClearQuery;
        private Label lblQueryHint;

        // Splitter
        private SplitContainer splitMain;

        // Result tabs
        private TabControl tabResults;
        private TabPage tabPageResults;
        private TabPage tabPageMessages;
        private DataGridView grid;
        private RichTextBox rtbMessages;

        // Bottom toolbar
        private Panel pnlQueryToolbar2;
        private Button btnClearResults;
        private Button btnExportCsv;

        // Status bar
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblStatus;
        private ToolStripStatusLabel lblRowCount;
        private ToolStripStatusLabel lblExecTime;
        private ToolStripProgressBar progressBar;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            // ── Toolbar (top) ────────────────────────────────────────────────────
            pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(37, 37, 38),
                Padding = new Padding(6, 6, 6, 6)
            };

            picStatus = new Panel
            {
                Width = 14,
                Height = 14,
                BackColor = Color.Gray,
                Location = new Point(10, 15),
            };
            picStatus.Region = System.Drawing.Region.FromHrgn(
                CreateRoundRectRgn(0, 0, 14, 14, 14, 14));

            btnConnect = new Button
            {
                Text = "Connect",
                Width = 90,
                Height = 30,
                Location = new Point(30, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };
            btnConnect.FlatAppearance.BorderSize = 0;
            btnConnect.Click += btnConnect_Click;

            btnDisconnect = new Button
            {
                Text = "Disconnect",
                Width = 90,
                Height = 30,
                Location = new Point(128, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                Enabled = false
            };
            btnDisconnect.FlatAppearance.BorderSize = 0;
            btnDisconnect.Click += btnDisconnect_Click;

            lblConnInfo = new Label
            {
                Text = "Not connected",
                ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 9f),
                AutoSize = true,
                Location = new Point(228, 13)
            };

            pnlToolbar.Controls.AddRange(new Control[] { picStatus, btnConnect, btnDisconnect, lblConnInfo });

            // ── Main SplitContainer ──────────────────────────────────────────────
            splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 220,
                Panel1MinSize = 80,
                Panel2MinSize = 80,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            // ── Query panel toolbar ──────────────────────────────────────────────
            pnlQueryToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(45, 45, 48),
                Padding = new Padding(4, 4, 4, 4)
            };

            btnExecute = new Button
            {
                Text = "▶  Execute (Ctrl+Enter)",
                Height = 28,
                Width = 175,
                Location = new Point(4, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                Enabled = false
            };
            btnExecute.FlatAppearance.BorderSize = 0;
            btnExecute.Click += btnExecute_Click;

            btnExecuteNonQuery = new Button
            {
                Text = "⚡ Non-Query",
                Height = 28,
                Width = 110,
                Location = new Point(185, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(80, 60, 20),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                Enabled = false
            };
            btnExecuteNonQuery.FlatAppearance.BorderSize = 0;
            btnExecuteNonQuery.Click += btnExecuteNonQuery_Click;

            btnClearQuery = new Button
            {
                Text = "Clear",
                Height = 28,
                Width = 65,
                Location = new Point(301, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            btnClearQuery.FlatAppearance.BorderSize = 0;
            btnClearQuery.Click += btnClearQuery_Click;

            lblQueryHint = new Label
            {
                Text = "Tip: Select text to run partial query",
                ForeColor = Color.FromArgb(130, 130, 130),
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                AutoSize = true,
                Location = new Point(375, 10)
            };

            pnlQueryToolbar.Controls.AddRange(new Control[]
            {
                btnExecute, btnExecuteNonQuery, btnClearQuery, lblQueryHint
            });

            // ── Query editor ─────────────────────────────────────────────────────
            rtbQuery = new RichTextBox
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
            rtbQuery.KeyDown += rtbQuery_KeyDown;

            splitMain.Panel1.Controls.Add(rtbQuery);
            splitMain.Panel1.Controls.Add(pnlQueryToolbar);

            // ── Results area ─────────────────────────────────────────────────────
            pnlQueryToolbar2 = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(45, 45, 48),
                Padding = new Padding(4, 4, 4, 4)
            };

            btnClearResults = new Button
            {
                Text = "Clear",
                Height = 28,
                Width = 65,
                Location = new Point(4, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            btnClearResults.FlatAppearance.BorderSize = 0;
            btnClearResults.Click += btnClearResults_Click;

            btnExportCsv = new Button
            {
                Text = "Export CSV",
                Height = 28,
                Width = 90,
                Location = new Point(76, 5),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            btnExportCsv.FlatAppearance.BorderSize = 0;
            btnExportCsv.Click += btnExportCsv_Click;

            pnlQueryToolbar2.Controls.AddRange(new Control[] { btnClearResults, btnExportCsv });

            // DataGridView
            grid = new DataGridView
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
            grid.DefaultCellStyle.BackColor = Color.FromArgb(30, 30, 30);
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(212, 212, 212);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 122, 204);
            grid.DefaultCellStyle.SelectionForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(200, 200, 200);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(37, 37, 38);
            grid.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 48);
            grid.EnableHeadersVisualStyles = false;

            // Messages box
            rtbMessages = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 10f),
                BackColor = Color.FromArgb(20, 20, 20),
                ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            // Tabs
            tabPageResults = new TabPage("Results");
            tabPageResults.Controls.Add(grid);

            tabPageMessages = new TabPage("Messages");
            tabPageMessages.Controls.Add(rtbMessages);

            tabResults = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f)
            };
            tabResults.TabPages.AddRange(new[] { tabPageResults, tabPageMessages });

            splitMain.Panel2.Controls.Add(tabResults);
            splitMain.Panel2.Controls.Add(pnlQueryToolbar2);

            // ── Status bar ───────────────────────────────────────────────────────
            statusStrip = new StatusStrip { BackColor = Color.FromArgb(0, 122, 204) };

            lblStatus = new ToolStripStatusLabel("Ready")
            {
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            lblRowCount = new ToolStripStatusLabel
            {
                ForeColor = Color.White,
                Spring = true,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9f)
            };
            lblExecTime = new ToolStripStatusLabel
            {
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            progressBar = new ToolStripProgressBar
            {
                Width = 100,
                Visible = false,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            statusStrip.Items.AddRange(new ToolStripItem[]
            {
                lblStatus, lblRowCount, lblExecTime, progressBar
            });

            // ── Assemble form ────────────────────────────────────────────────────
            SuspendLayout();
            Text = "SQL Query Tool";
            Size = new Size(1100, 720);
            MinimumSize = new Size(700, 500);
            BackColor = Color.FromArgb(30, 30, 30);
            StartPosition = FormStartPosition.CenterScreen;

            Controls.Add(splitMain);
            Controls.Add(pnlToolbar);
            Controls.Add(statusStrip);
            ResumeLayout(false);
            PerformLayout();
        }

        [System.Runtime.InteropServices.DllImport("Gdi32.dll")]
        private static extern System.IntPtr CreateRoundRectRgn(
            int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
            int nWidthEllipse, int nHeightEllipse);
    }
}
