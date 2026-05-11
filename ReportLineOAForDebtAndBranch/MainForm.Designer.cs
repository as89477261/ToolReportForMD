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

        // Query toolbar
        private Panel pnlQueryToolbar;
        private Button btnExecute;
        private Button btnExecuteNonQuery;
        private Button btnNewTab;

        // Query tabs
        private TabControl tabQueries;

        // Outer splitter (left = tabs | right = column browser)
        private SplitContainer splitOuter;

        // Column browser
        private Panel pnlColumnBrowser;
        private Label lblColumnBrowserHeader;
        private TextBox txtTableSearch;
        private Label lblColumnBrowserStatus;
        private TreeView treeColumns;

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

            // ── Toolbar (top) ────────────────────────────────────────────────────
            pnlToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 44,
                BackColor = Color.FromArgb(37, 37, 38)
            };

            picStatus = new Panel
            {
                Width = 14, Height = 14,
                BackColor = Color.Gray,
                Location = new Point(10, 15)
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
                Text = "Not connected",
                ForeColor = Color.FromArgb(180, 180, 180),
                Font = new Font("Segoe UI", 9f),
                AutoSize = true, Location = new Point(228, 13)
            };

            pnlToolbar.Controls.AddRange(new Control[] { picStatus, btnConnect, btnDisconnect, lblConnInfo });

            // ── Query toolbar ────────────────────────────────────────────────────
            pnlQueryToolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
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

            var lblHint = new Label
            {
                Text = "Tip: Select text to run partial query  |  Ctrl+W = close tab",
                ForeColor = Color.FromArgb(110, 110, 110),
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                AutoSize = true, Location = new Point(402, 11)
            };

            pnlQueryToolbar.Controls.AddRange(new Control[]
            {
                btnExecute, btnExecuteNonQuery, btnNewTab, lblHint
            });

            // ── Query tabs ───────────────────────────────────────────────────────
            tabQueries = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f)
            };
            tabQueries.MouseClick += tabQueries_MouseClick;

            // ── Column browser ───────────────────────────────────────────────────
            pnlColumnBrowser = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30)
            };

            lblColumnBrowserHeader = new Label
            {
                Text = "Column Browser",
                Dock = DockStyle.Top, Height = 28,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(200, 200, 200),
                BackColor = Color.FromArgb(45, 45, 48)
            };

            txtTableSearch = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font("Segoe UI", 9f),
                BackColor = Color.FromArgb(60, 60, 60),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                PlaceholderText = "Type table name + Enter to search..."
            };
            txtTableSearch.KeyDown += txtTableSearch_KeyDown;

            lblColumnBrowserStatus = new Label
            {
                Dock = DockStyle.Top, Height = 20,
                Text = "Execute a query to see columns",
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = Color.FromArgb(120, 120, 120),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                BackColor = Color.FromArgb(30, 30, 30)
            };

            treeColumns = new TreeView
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.FromArgb(212, 212, 212),
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 9f),
                ShowLines = true, ShowPlusMinus = true,
                FullRowSelect = true, HideSelection = false
            };
            treeColumns.NodeMouseDoubleClick += treeColumns_NodeMouseDoubleClick;

            pnlColumnBrowser.Controls.Add(treeColumns);
            pnlColumnBrowser.Controls.Add(lblColumnBrowserStatus);
            pnlColumnBrowser.Controls.Add(txtTableSearch);
            pnlColumnBrowser.Controls.Add(lblColumnBrowserHeader);

            // ── Left panel ───────────────────────────────────────────────────────
            var pnlLeft = new Panel { Dock = DockStyle.Fill };
            pnlLeft.Controls.Add(tabQueries);
            pnlLeft.Controls.Add(pnlQueryToolbar);

            // ── Outer splitter ───────────────────────────────────────────────────
            splitOuter = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = Color.FromArgb(30, 30, 30)
            };
            splitOuter.Panel1.Controls.Add(pnlLeft);
            splitOuter.Panel2.Controls.Add(pnlColumnBrowser);

            // ── Status bar ───────────────────────────────────────────────────────
            statusStrip = new StatusStrip { BackColor = Color.FromArgb(0, 122, 204) };
            lblStatus = new ToolStripStatusLabel("Ready")
            {
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                Spring = true,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            progressBar = new ToolStripProgressBar
            {
                Width = 120, Visible = false,
                Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30
            };
            statusStrip.Items.AddRange(new ToolStripItem[] { lblStatus, progressBar });

            // ── Assemble form ────────────────────────────────────────────────────
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
