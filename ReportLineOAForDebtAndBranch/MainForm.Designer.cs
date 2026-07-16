using System.Drawing;
using System.Windows.Forms;

namespace ReportLineOAForDebtAndBranch
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        // Top toolbar
        private Panel pnlToolbar;
        private Panel picStatus;
        private Label lblConnInfo;
        private Button btnConnect;
        private Button btnDisconnect;
        private Button btnCompareDb;

        // Query toolbar
        private Panel pnlQueryToolbar;
        private Button btnExecute;
        private Button btnExecuteNonQuery;
        private Button btnNewTab;
        private NumericUpDown numTimeout;

        // Query tabs
        private TabControl tabQueries;

        // Outer splitter
        private SplitContainer splitOuter;

        // History panel (right side)
        private Panel pnlHistory;
        private Label lblHistoryHeader;
        private Panel pnlHistoryTools;
        private TextBox txtHistoryFilter;
        private Button btnClearHistory;
        private ListView lvHistory;

        // Status bar
        private StatusStrip statusStrip;
        private ToolStripStatusLabel lblStatus;
        private ToolStripProgressBar progressBar;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            // ── Top toolbar ──────────────────────────────────────────────────────
            pnlToolbar = new Panel
            {
                Dock = DockStyle.Top, Height = 44,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            picStatus = new Panel
            {
                Width = 14, Height = 14,
                BackColor = Color.Gray, Location = new Point(10, 15)
            };
            picStatus.Region = System.Drawing.Region.FromHrgn(
                CreateRoundRectRgn(0, 0, 14, 14, 14, 14));

            btnConnect = new Button
            {
                Text = "Connect", Width = 90, Height = 30, Location = new Point(30, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f)
            };
            btnConnect.FlatAppearance.BorderSize = 0;
            btnConnect.Click += btnConnect_Click;

            btnDisconnect = new Button
            {
                Text = "Disconnect", Width = 90, Height = 30, Location = new Point(128, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f), Enabled = false
            };
            btnDisconnect.FlatAppearance.BorderSize = 0;
            btnDisconnect.Click += btnDisconnect_Click;

            lblConnInfo = new Label
            {
                Text = "Not connected", AutoSize = true, Location = new Point(228, 13),
                ForeColor = Color.FromArgb(180, 180, 180), Font = new Font("Segoe UI", 9f)
            };

            btnCompareDb = new Button
            {
                Text = "⇄ Compare DB", Width = 120, Height = 30, Location = new Point(500, 7),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(40, 70, 100), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f), Enabled = false
            };
            btnCompareDb.FlatAppearance.BorderSize = 0;
            btnCompareDb.Click += btnCompareDb_Click;

            pnlToolbar.Controls.AddRange(new Control[] { picStatus, btnConnect, btnDisconnect, lblConnInfo, btnCompareDb });

            // ── Query toolbar ────────────────────────────────────────────────────
            pnlQueryToolbar = new Panel
            {
                Dock = DockStyle.Top, Height = 38,
                BackColor = Color.FromArgb(45, 45, 48)
            };

            btnExecute = new Button
            {
                Text = "▶  Execute (Ctrl+Enter)", Height = 28, Width = 175,
                Location = new Point(4, 5), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f), Enabled = false
            };
            btnExecute.FlatAppearance.BorderSize = 0;
            btnExecute.Click += btnExecute_Click;

            btnExecuteNonQuery = new Button
            {
                Text = "⚡ Non-Query", Height = 28, Width = 110,
                Location = new Point(185, 5), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(80, 60, 20), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f), Enabled = false
            };
            btnExecuteNonQuery.FlatAppearance.BorderSize = 0;
            btnExecuteNonQuery.Click += btnExecuteNonQuery_Click;

            btnNewTab = new Button
            {
                Text = "+ New Tab", Height = 28, Width = 90,
                Location = new Point(305, 5), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 58), ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 9f)
            };
            btnNewTab.FlatAppearance.BorderSize = 0;
            btnNewTab.Click += btnNewTab_Click;

            // Timeout setting
            var lblTimeout = new Label
            {
                Text = "Timeout:", AutoSize = true, Location = new Point(406, 12),
                ForeColor = Color.FromArgb(160, 160, 160), Font = new Font("Segoe UI", 9f)
            };

            numTimeout = new NumericUpDown
            {
                Location = new Point(462, 6), Width = 62, Height = 24,
                Minimum = 10, Maximum = 7200, Value = 300, Increment = 30,
                BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f), BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center
            };

            var lblSec = new Label
            {
                Text = "s", AutoSize = true, Location = new Point(528, 12),
                ForeColor = Color.FromArgb(160, 160, 160), Font = new Font("Segoe UI", 9f)
            };

            var lblHint = new Label
            {
                Text = "Ctrl+T = new tab  |  Ctrl+W = close",
                ForeColor = Color.FromArgb(90, 90, 90),
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                AutoSize = true, Location = new Point(548, 12)
            };

            pnlQueryToolbar.Controls.AddRange(new Control[]
            {
                btnExecute, btnExecuteNonQuery, btnNewTab,
                lblTimeout, numTimeout, lblSec, lblHint
            });

            // ── Query tabs ───────────────────────────────────────────────────────
            tabQueries = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f)
            };
            tabQueries.MouseClick += tabQueries_MouseClick;

            var pnlLeft = new Panel { Dock = DockStyle.Fill };
            pnlLeft.Controls.Add(tabQueries);
            pnlLeft.Controls.Add(pnlQueryToolbar);

            // ── History panel (right) ────────────────────────────────────────────
            pnlHistory = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            lblHistoryHeader = new Label
            {
                Dock = DockStyle.Top, Height = 28,
                Text = "Query History",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(200, 200, 200),
                BackColor = Color.FromArgb(45, 45, 48)
            };

            pnlHistoryTools = new Panel
            {
                Dock = DockStyle.Top, Height = 30,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            txtHistoryFilter = new TextBox
            {
                Location = new Point(4, 4), Height = 22, Width = 150,
                Font = new Font("Segoe UI", 9f),
                BackColor = Color.FromArgb(60, 60, 60), ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "Filter..."
            };
            txtHistoryFilter.TextChanged += txtHistoryFilter_TextChanged;

            btnClearHistory = new Button
            {
                Text = "Clear History", Height = 22, Width = 95,
                Location = new Point(160, 4), FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(70, 70, 72), ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 8f)
            };
            btnClearHistory.FlatAppearance.BorderSize = 0;
            btnClearHistory.Click += btnClearHistory_Click;

            pnlHistoryTools.Controls.AddRange(new Control[] { txtHistoryFilter, btnClearHistory });

            // ── History ListView ─────────────────────────────────────────────────
            lvHistory = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 8.5f),
                OwnerDraw = true,
                ShowItemToolTips = true
            };
            lvHistory.Columns.Add("Query", 185);
            lvHistory.Columns.Add("Time", 55);
            lvHistory.DrawColumnHeader  += LvHistory_DrawColumnHeader;
            lvHistory.DrawItem          += LvHistory_DrawItem;
            lvHistory.DrawSubItem       += LvHistory_DrawSubItem;
            lvHistory.MouseClick        += lvHistory_MouseClick;
            lvHistory.MouseDoubleClick  += lvHistory_MouseDoubleClick;

            pnlHistory.Controls.Add(lvHistory);
            pnlHistory.Controls.Add(pnlHistoryTools);
            pnlHistory.Controls.Add(lblHistoryHeader);

            // ── Outer splitter ───────────────────────────────────────────────────
            splitOuter = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            splitOuter.Panel1.Controls.Add(pnlLeft);
            splitOuter.Panel2.Controls.Add(pnlHistory);

            // ── Status bar ───────────────────────────────────────────────────────
            statusStrip = new StatusStrip { BackColor = Color.FromArgb(0, 122, 204) };
            lblStatus = new ToolStripStatusLabel("Ready")
            {
                ForeColor = Color.White, Font = new Font("Segoe UI", 9f),
                Spring = true, TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            progressBar = new ToolStripProgressBar
            {
                Width = 120, Visible = false,
                Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30
            };
            statusStrip.Items.AddRange(new ToolStripItem[] { lblStatus, progressBar });

            // ── Assemble ─────────────────────────────────────────────────────────
            SuspendLayout();
            Text = "SQL Query Tool";
            Size = new Size(1280, 720);
            MinimumSize = new Size(800, 500);
            BackColor = Color.FromArgb(30, 30, 30);
            StartPosition = FormStartPosition.CenterScreen;

            Controls.Add(splitOuter);
            Controls.Add(pnlQueryToolbar);
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
