using ExcelForge.Converters;
using ExcelForge.Updater;
using ExcelForge.UI;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.EventLog;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelForge;

internal static class Program
{
    internal static ILoggerFactory LoggerFactory { get; private set; } = null!;

    [STAThread]
    static void Main(string[] args)
    {

        LoggerFactory = Microsoft.Extensions.Logging.LoggerFactory.Create(builder =>
        {
            builder.AddEventLog(new EventLogSettings
            {
                SourceName = AppConstants.EventLogSource,
                LogName    = AppConstants.EventLogName,
            });
#if DEBUG
            builder.AddConsole();
#endif
            builder.SetMinimumLevel(LogLevel.Information);
        });

        var logger = LoggerFactory.CreateLogger("ExcelForge.Program");
        logger.LogInformation("{App} {Version} started", AppConstants.AppName, AppConstants.VersionDisplay);

        if (args.Length > 0)
        {
            Environment.Exit(RunCli(args, logger).GetAwaiter().GetResult());
            return;
        }

        ApplicationConfiguration.Initialize();

        _ = Task.Run(() => CheckForUpdatesBackground(logger));

        Application.Run(new MainForm());
    }

    private static async Task CheckForUpdatesBackground(ILogger logger)
    {

        await Task.Delay(TimeSpan.FromSeconds(3));

        var mainForm = Application.OpenForms.OfType<MainForm>().FirstOrDefault();
        if (mainForm is null || mainForm.IsDisposed) return;

        bool upToDate  = false;
        bool netError  = false;

        try
        {
            using var updater = new GitHubUpdater(
                LoggerFactory.CreateLogger<GitHubUpdater>());

            bool updated = await updater.CheckAndApplyAsync();
            upToDate = !updated;
        }
        catch (HttpRequestException ex)
        {
            netError = true;
            logger.LogWarning(ex, "Background update check failed — network error");
        }
        catch (TaskCanceledException ex)
        {
            netError = true;
            logger.LogWarning(ex, "Background update check failed — request timed out");
        }
        catch (Exception ex)
        {
            netError = true;
            logger.LogWarning(ex, "Background update check failed");
        }

        if (mainForm.IsDisposed) return;

        if (upToDate)
            mainForm.Invoke(() => ShowUpdateToast(mainForm, success: true));
        else if (netError)
            mainForm.Invoke(() => ShowUpdateToast(mainForm, success: false));
    }

    private static void ShowUpdateToast(Form owner, bool success)
    {
        var toast = new Form
        {
            FormBorderStyle = FormBorderStyle.None,
            StartPosition   = FormStartPosition.Manual,
            BackColor       = success
                ? Color.FromArgb(30, 73, 125)
                : Color.FromArgb(180, 90, 20),
            Size            = new Size(280, 48),
            TopMost         = true,
            ShowInTaskbar   = false,
            Opacity         = 0.92,
        };

        var lbl = new Label
        {
            Text      = success
                ? "\u2714  ExcelForge is up to date"
                : "\u26a0  Update check failed",
            ForeColor = Color.White,
            Font      = new Font("Segoe UI", 10F),
            TextAlign = ContentAlignment.MiddleCenter,
            Dock      = DockStyle.Fill,
        };
        toast.Controls.Add(lbl);

        toast.Location = new Point(
            owner.Right  - toast.Width  - 16,
            owner.Bottom - toast.Height - 48);

        var timer = new System.Windows.Forms.Timer { Interval = 5000 };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            timer.Dispose();
            if (!toast.IsDisposed) toast.Close();
        };

        toast.Shown += (_, _) => timer.Start();
        toast.Show(owner);
    }

    private static async Task<int> RunCli(string[] args, ILogger logger)
    {

        if (args[0] is "-h" or "--help" or "help")
        {
            PrintHelp();
            return 0;
        }

        if (args[0] is "-v" or "--version" or "version")
        {
            Console.WriteLine($"{AppConstants.AppName} {AppConstants.VersionDisplay}");
            return 0;
        }

        if (args[0] == "convert" && args.Length >= 3)
        {
            string src    = args[1];
            string dst    = args[2];
            string format = args.Length >= 5 && args[3] == "--format" ? args[4] : "xlsx";

            if (!File.Exists(src))
            {
                Console.Error.WriteLine($"[error] File not found: {src}");
                return 2;
            }

            try
            {
                IConverter converter = format.ToLowerInvariant() switch
                {
                    "csv"  => new CsvExporter(LoggerFactory.CreateLogger<CsvExporter>()),
                    "json" => new JsonExporter(LoggerFactory.CreateLogger<JsonExporter>()),
                    _      => new XlsxConverter(LoggerFactory.CreateLogger<XlsxConverter>()),
                };

                Console.WriteLine($"[info] Converting {Path.GetFileName(src)} → {format.ToUpper()} ...");
                await converter.ConvertAsync(src, dst);
                Console.WriteLine($"[ok]   Output written to {dst}");
                return 0;
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Conversion failed");
                Console.Error.WriteLine($"[error] {ex.Message}");
                return 1;
            }
        }

        if (args[0] == "batch" && args.Length >= 2)
        {
            return await RunBatch(args, logger);
        }

        if (args[0] == "update")
        {
            Console.WriteLine("[info] Checking for updates...");
            using var updater = new GitHubUpdater(LoggerFactory.CreateLogger<GitHubUpdater>());
            bool updated = await updater.CheckAndApplyAsync(printToConsole: true);
            Console.WriteLine(updated ? "[ok]   Update applied." : "[ok]   Already up to date.");
            return 0;
        }

        Console.Error.WriteLine($"[error] Unknown command '{args[0]}'. Run 'excelforge --help'.");
        return 1;
    }

    private static async Task<int> RunBatch(string[] args, ILogger logger)
    {
        string dir     = args[1];
        string format  = "xlsx";
        string? outDir = null;

        for (int i = 2; i < args.Length - 1; i++)
        {
            if (args[i] == "--format") format = args[i + 1];
            if (args[i] == "--out")    outDir = args[i + 1];
        }

        if (!Directory.Exists(dir))
        {
            Console.Error.WriteLine($"[error] Directory not found: {dir}");
            return 2;
        }

        outDir ??= Path.Combine(dir, "output");
        Directory.CreateDirectory(outDir);

        var files = Directory.GetFiles(dir, "*.xls*", SearchOption.TopDirectoryOnly);
        Console.WriteLine($"[info] Found {files.Length} file(s). Output → {outDir}");

        IConverter converter = format.ToLowerInvariant() switch
        {
            "csv"  => new CsvExporter(LoggerFactory.CreateLogger<CsvExporter>()),
            "json" => new JsonExporter(LoggerFactory.CreateLogger<JsonExporter>()),
            _      => new XlsxConverter(LoggerFactory.CreateLogger<XlsxConverter>()),
        };

        int ok = 0, fail = 0;
        foreach (var file in files)
        {
            string dst = Path.Combine(outDir,
                Path.GetFileNameWithoutExtension(file) + "." + format);
            try
            {
                await converter.ConvertAsync(file, dst);
                Console.WriteLine($"  [ok]  {Path.GetFileName(file)}");
                ok++;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"  [err] {Path.GetFileName(file)}: {ex.Message}");
                logger.LogError(ex, "Batch conversion failed for {File}", file);
                fail++;
            }
        }

        Console.WriteLine($"[done] {ok} converted, {fail} failed.");
        return fail > 0 ? 1 : 0;
    }

    private static void PrintHelp()
    {
        Console.WriteLine($"""
            ExcelForge {AppConstants.VersionDisplay} — Excel automation utility
            https://github.com/BaadHackor/excelforge

            USAGE:
              excelforge <command> [options]

            COMMANDS:
              convert <src> <dst> [--format xlsx|csv|json]
                  Convert a single Excel file.

              batch <dir> [--format xlsx|csv|json] [--out <outdir>]
                  Batch-convert all .xls/.xlsx files in a directory.

              update
                  Check for updates and apply if available.

              --version, -v   Print version and exit.
              --help,    -h   Show this help.

            EXAMPLES:
              excelforge convert report.xls report.xlsx
              excelforge convert data.xlsx export.csv --format csv
              excelforge batch ./invoices --format json --out ./json_out
            """);
    }
}
