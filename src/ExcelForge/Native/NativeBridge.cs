using Microsoft.Extensions.Logging;
using System.Runtime.InteropServices;

namespace ExcelForge.Native;

internal sealed class NativeBridge : IDisposable
{

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate bool ExcelNative_InitDelegate(IntPtr hParent);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int ExcelNative_ConvertDelegate(
        [MarshalAs(UnmanagedType.LPWStr)] string lpSrc,
        [MarshalAs(UnmanagedType.LPWStr)] string lpDst,
        uint flags);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate void ExcelNative_FreeDelegate();

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr LoadLibrary(string lpLibFileName);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FreeLibrary(IntPtr hModule);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
    private static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

    private readonly ILogger  _logger;
    private readonly IntPtr   _hModule;
    private bool              _initialized;
    private bool              _disposed;

    private readonly ExcelNative_InitDelegate?    _init;
    private readonly ExcelNative_ConvertDelegate? _convert;
    private readonly ExcelNative_FreeDelegate?    _free;

    public bool IsAvailable { get; }

    public NativeBridge(ILogger logger)
    {
        _logger = logger;

        string dllPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            AppConstants.NativeDllName);

        if (!File.Exists(dllPath))
        {
            _logger.LogDebug("NativeBridge: {Dll} not found at {Path}", AppConstants.NativeDllName, dllPath);
            IsAvailable = false;
            return;
        }

        _hModule = LoadLibrary(dllPath);
        if (_hModule == IntPtr.Zero)
        {
            int err = Marshal.GetLastWin32Error();
            _logger.LogWarning(
                "NativeBridge: LoadLibrary failed (Win32 error 0x{Err:X4})", err);
            IsAvailable = false;
            return;
        }

        _init    = ResolveExport<ExcelNative_InitDelegate>   ("ExcelNative_Init");
        _convert = ResolveExport<ExcelNative_ConvertDelegate>("ExcelNative_Convert");
        _free    = ResolveExport<ExcelNative_FreeDelegate>   ("ExcelNative_Free");

        IsAvailable = _init is not null && _convert is not null && _free is not null;

        if (IsAvailable)
            _logger.LogInformation("NativeBridge: loaded {Dll}", AppConstants.NativeDllName);
        else
            _logger.LogWarning("NativeBridge: one or more exports missing — bridge disabled");
    }

    public void Init(IntPtr hwndParent = default)
    {
        ThrowIfNotAvailable();
        if (_initialized) return;

        bool ok = _init!(hwndParent);
        if (!ok)
            throw new InvalidOperationException("ExcelNative_Init returned FALSE.");

        _initialized = true;
        _logger.LogDebug("NativeBridge: Init() OK");
    }

    public int Convert(string src, string dst, uint flags = 0)
    {
        ThrowIfNotAvailable();
        if (!_initialized) Init();

        int result = _convert!(src, dst, flags);
        _logger.LogDebug("NativeBridge: Convert({Src}) → {Result}", src, result);
        return result;
    }

    private T? ResolveExport<T>(string name) where T : Delegate
    {
        IntPtr ptr = GetProcAddress(_hModule, name);
        if (ptr == IntPtr.Zero)
        {
            _logger.LogDebug("NativeBridge: export '{Name}' not found", name);
            return null;
        }
        return Marshal.GetDelegateForFunctionPointer<T>(ptr);
    }

    private void ThrowIfNotAvailable()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        if (!IsAvailable)
            throw new InvalidOperationException(
                $"{AppConstants.NativeDllName} is not available on this system.");
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_initialized && _free is not null)
        {
            try { _free(); }
            catch (Exception ex) { _logger.LogWarning(ex, "NativeBridge: Free() threw"); }
        }

        if (_hModule != IntPtr.Zero)
            FreeLibrary(_hModule);

        _logger.LogDebug("NativeBridge: unloaded");
    }
}
