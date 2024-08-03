namespace ExcelForge.UI;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    private Panel        pnlTop;
    private Label        lblTitle;
    private Label        lblSubtitle;
    private Panel        pnlDrop;
    private Label        lblDrop;
    private Label        lblFileName;
    private Button       btnBrowse;
    private GroupBox     grpOptions;
    private Label        lblFormat;
    private ComboBox     cmbFormat;
    private Button       btnConvert;
    private Button       btnClear;
    private Panel        pnlStatus;
    private Label        statusLabel;
    private ProgressBar  progressBar;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
            components.Dispose();
        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();

        SuspendLayout();
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode       = AutoScaleMode.Font;
        ClientSize          = new Size(560, 440);
        FormBorderStyle     = FormBorderStyle.FixedSingle;
        MaximizeBox         = false;
        StartPosition       = FormStartPosition.CenterScreen;
        Font                = new Font("Segoe UI", 9F);
        BackColor           = Color.WhiteSmoke;

        pnlTop = new Panel
        {
            Dock      = DockStyle.Top,
            Height    = 68,
            BackColor = Color.FromArgb(30, 73, 125),
            Padding   = new Padding(16, 0, 16, 0),
        };

        lblTitle = new Label
        {
            Text      = "ExcelForge",
            Font      = new Font("Segoe UI", 18F, FontStyle.Bold),
            ForeColor = Color.White,
            AutoSize  = true,
            Location  = new Point(16, 8),
        };

        lblSubtitle = new Label
        {
            Text      = "Excel automation utility",
            Font      = new Font("Segoe UI", 9F, FontStyle.Italic),
            ForeColor = Color.FromArgb(180, 210, 240),
            AutoSize  = true,
            Location  = new Point(20, 44),
        };

        pnlTop.Controls.AddRange(new Control[] { lblTitle, lblSubtitle });

        pnlDrop = new Panel
        {
            Left        = 24,
            Top         = 88,
            Width       = 512,
            Height      = 110,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor   = Color.White,
            Cursor      = Cursors.Hand,
        };

        lblDrop = new Label
        {
            Text      = "Drag & drop an Excel file here",
            Font      = new Font("Segoe UI", 11F),
            ForeColor = Color.Gray,
            TextAlign = ContentAlignment.MiddleCenter,
            Dock      = DockStyle.Fill,
        };

        lblFileName = new Label
        {
            Text      = "No file selected",
            ForeColor = SystemColors.GrayText,
            Font      = new Font("Segoe UI", 8.5F),
            TextAlign = ContentAlignment.MiddleCenter,
            Dock      = DockStyle.Bottom,
            Height    = 24,
        };

        pnlDrop.Controls.AddRange(new Control[] { lblDrop, lblFileName });

        btnBrowse = new Button
        {
            Text      = "Browse…",
            Left      = 24,
            Top       = 210,
            Width     = 100,
            Height    = 32,
            FlatStyle = FlatStyle.System,
        };
        btnBrowse.Click += btnBrowse_Click;

        grpOptions = new GroupBox
        {
            Text    = "Output format",
            Left    = 24,
            Top     = 256,
            Width   = 512,
            Height  = 64,
        };

        lblFormat = new Label
        {
            Text     = "Format:",
            Left     = 12,
            Top      = 28,
            AutoSize = true,
        };

        cmbFormat = new ComboBox
        {
            Left          = 72,
            Top           = 24,
            Width         = 120,
            DropDownStyle = ComboBoxStyle.DropDownList,
        };
        cmbFormat.Items.AddRange(new object[] { "XLSX", "CSV", "JSON" });
        cmbFormat.SelectedIndex = 0;

        grpOptions.Controls.AddRange(new Control[] { lblFormat, cmbFormat });

        btnConvert = new Button
        {
            Text      = "Convert",
            Left      = 336,
            Top       = 336,
            Width     = 100,
            Height    = 36,
            Enabled   = false,
            FlatStyle = FlatStyle.System,
            Font      = new Font("Segoe UI", 9.5F, FontStyle.Bold),
        };
        btnConvert.Click += btnConvert_Click;

        btnClear = new Button
        {
            Text      = "Clear",
            Left      = 444,
            Top       = 336,
            Width     = 92,
            Height    = 36,
            FlatStyle = FlatStyle.System,
        };
        btnClear.Click += btnClear_Click;

        pnlStatus = new Panel
        {
            Dock      = DockStyle.Bottom,
            Height    = 36,
            BackColor = Color.FromArgb(240, 240, 240),
            Padding   = new Padding(8, 0, 8, 0),
        };

        statusLabel = new Label
        {
            Text      = "Ready",
            AutoSize  = false,
            Dock      = DockStyle.Left,
            Width     = 350,
            TextAlign = ContentAlignment.MiddleLeft,
            Font      = new Font("Segoe UI", 8.5F),
        };

        progressBar = new ProgressBar
        {
            Dock    = DockStyle.Right,
            Width   = 140,
            Height  = 20,
            Margin  = new Padding(0, 8, 8, 8),
            Value   = 0,
            Style   = ProgressBarStyle.Blocks,
        };

        pnlStatus.Controls.AddRange(new Control[] { statusLabel, progressBar });

        Controls.AddRange(new Control[]
        {
            pnlTop,
            pnlDrop,
            btnBrowse,
            grpOptions,
            btnConvert,
            btnClear,
            pnlStatus,
        });

        ResumeLayout(false);
        PerformLayout();
    }
}
