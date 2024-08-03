using ExcelForge.Converters;
using Microsoft.Extensions.Logging;

namespace ExcelForge.UI;

internal sealed partial class MainForm : Form
{
    private readonly ILogger<MainForm> _logger;

    private string? _selectedFile;
    private bool    _busy;

    public MainForm()
    {
        _logger = Program.LoggerFactory.CreateLogger<MainForm>();
        InitializeComponent();

        AllowDrop = true;
        DragEnter += OnDragEnter;
        DragDrop  += OnDragDrop;

        UpdateTitle();
        SetStatus($"Ready — {AppConstants.VersionDisplay}", StatusType.Info);
    }

    private void btnBrowse_Click(object sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title       = "Select an Excel file",
            Filter      = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*",
            Multiselect = false,
        };

        if (dlg.ShowDialog(this) == DialogResult.OK)
            LoadFile(dlg.FileName);
    }

    private async void btnConvert_Click(object sender, EventArgs e)
    {
        if (_busy || _selectedFile is null) return;

        string format = cmbFormat.SelectedItem?.ToString()?.ToLowerInvariant() ?? "xlsx";
        string ext    = format == "json" ? ".json" : format == "csv" ? ".csv" : ".xlsx";
        string dstDefault = Path.ChangeExtension(_selectedFile, ext);

        using var dlg = new SaveFileDialog
        {
            Title            = "Save output as",
            FileName         = Path.GetFileName(dstDefault),
            InitialDirectory = Path.GetDirectoryName(_selectedFile),
            Filter           = format switch
            {
                "csv"  => "CSV files (*.csv)|*.csv",
                "json" => "JSON files (*.json)|*.json",
                _      => "Excel files (*.xlsx)|*.xlsx",
            },
        };

        if (dlg.ShowDialog(this) != DialogResult.OK) return;

        await RunConversionAsync(_selectedFile, dlg.FileName, format);
    }

    private void btnClear_Click(object sender, EventArgs e)
    {
        _selectedFile = null;
        lblFileName.Text = "No file selected";
        lblFileName.ForeColor = SystemColors.GrayText;
        btnConvert.Enabled = false;
        SetStatus($"Ready — {AppConstants.VersionDisplay}", StatusType.Info);
    }

    private void OnDragEnter(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
            e.Effect = DragDropEffects.Copy;
    }

    private void OnDragDrop(object? sender, DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
            LoadFile(files[0]);
    }

    private void LoadFile(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is not (".xls" or ".xlsx"))
        {
            SetStatus($"Unsupported file type: {ext}", StatusType.Warning);
            return;
        }

        _selectedFile = path;
        lblFileName.Text      = Path.GetFileName(path);
        lblFileName.ForeColor = SystemColors.ControlText;
        btnConvert.Enabled    = true;

        long sz = new FileInfo(path).Length;
        SetStatus($"Loaded: {Path.GetFileName(path)} ({sz / 1024:N0} KB)", StatusType.Info);
        _logger.LogInformation("File selected: {Path}", path);
    }

    private async Task RunConversionAsync(string src, string dst, string format)
    {
        _busy = true;
        SetControlsEnabled(false);
        progressBar.Style = ProgressBarStyle.Marquee;

        SetStatus("Converting…", StatusType.Info);

        try
        {
            IConverter converter = format switch
            {
                "csv"  => new CsvExporter (Program.LoggerFactory.CreateLogger<CsvExporter>()),
                "json" => new JsonExporter(Program.LoggerFactory.CreateLogger<JsonExporter>()),
                _      => new XlsxConverter(Program.LoggerFactory.CreateLogger<XlsxConverter>()),
            };

            await converter.ConvertAsync(src, dst);

            SetStatus($"Done → {Path.GetFileName(dst)}", StatusType.Success);
            _logger.LogInformation("Conversion OK: {Dst}", dst);

            if (MessageBox.Show(
                    $"Conversion complete!\n\nOutput: {dst}\n\nOpen the output folder?",
                    AppConstants.AppName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information) == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start("explorer.exe",
                    $"/select,\"{dst}\"");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Conversion failed");
            SetStatus($"Error: {ex.Message}", StatusType.Error);
            MessageBox.Show($"Conversion failed:\n\n{ex.Message}",
                AppConstants.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _busy = false;
            SetControlsEnabled(true);
            progressBar.Style = ProgressBarStyle.Blocks;
            progressBar.Value = 0;
        }
    }

    private enum StatusType { Info, Success, Warning, Error }

    private void SetStatus(string msg, StatusType type)
    {
        statusLabel.Text = msg;
        statusLabel.ForeColor = type switch
        {
            StatusType.Success => Color.DarkGreen,
            StatusType.Warning => Color.DarkOrange,
            StatusType.Error   => Color.Crimson,
            _                  => SystemColors.ControlText,
        };
    }

    private void SetControlsEnabled(bool enabled)
    {
        btnBrowse.Enabled  = enabled;
        btnConvert.Enabled = enabled && _selectedFile is not null;
        btnClear.Enabled   = enabled;
        cmbFormat.Enabled  = enabled;
    }

    private void UpdateTitle() =>
        Text = $"{AppConstants.AppName} {AppConstants.VersionDisplay}";
}
