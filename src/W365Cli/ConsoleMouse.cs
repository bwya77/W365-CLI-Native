using System.Runtime.InteropServices;

namespace W365Cli;

/// <summary>
/// Windows-only mouse wheel support. The app runs on the alternate screen buffer (so there is no
/// scrollback to scroll into), which means the terminal's own wheel scrolling has nothing to
/// show. This enables console mouse input and turns wheel notches into Up/Down arrow presses so
/// every list and screen scrolls exactly like the keyboard does.
/// </summary>
internal static class ConsoleMouse
{
    private const int StdInputHandle = -10;
    private const uint EnableMouseInput = 0x0010;
    private const uint EnableQuickEdit = 0x0040;
    private const uint EnableExtendedFlags = 0x0080;
    private const ushort KeyEventType = 1;
    private const ushort MouseEventType = 2;
    private const uint MouseWheeled = 0x0004;

    private static IntPtr handle;
    private static uint originalMode;
    private static bool enabled;

    [StructLayout(LayoutKind.Sequential)]
    private struct InputRecord
    {
        public ushort EventType;
        public ushort Padding;
        public uint Word1; // key: bKeyDown | mouse: position
        public uint Word2; // key: repeat+virtual key | mouse: button state
        public uint Word3; // key: scan code+char | mouse: control keys
        public uint Word4; // key: control keys | mouse: event flags
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetConsoleMode(IntPtr handle, out uint mode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetConsoleMode(IntPtr handle, uint mode);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetNumberOfConsoleInputEvents(IntPtr handle, out uint count);

    [DllImport("kernel32.dll", EntryPoint = "PeekConsoleInputW", SetLastError = true)]
    private static extern bool PeekConsoleInput(IntPtr handle, out InputRecord record, uint length, out uint read);

    [DllImport("kernel32.dll", EntryPoint = "ReadConsoleInputW", SetLastError = true)]
    private static extern bool ReadConsoleInput(IntPtr handle, out InputRecord record, uint length, out uint read);

    public static void Enable()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || enabled)
        {
            return;
        }

        handle = GetStdHandle(StdInputHandle);
        if (handle == IntPtr.Zero || handle == new IntPtr(-1) || !GetConsoleMode(handle, out originalMode))
        {
            return;
        }

        var mode = (originalMode | EnableMouseInput | EnableExtendedFlags) & ~EnableQuickEdit;
        enabled = SetConsoleMode(handle, mode);
    }

    public static bool IsEnabled => enabled;

    public static void Disable()
    {
        if (enabled)
        {
            SetConsoleMode(handle, originalMode);
            enabled = false;
        }
    }

    /// <summary>
    /// Drains queued non-key input. Returns -1 / +1 for a wheel notch up / down, 0 when a real
    /// key press is next in the queue (left for Console.ReadKey), or null when the queue is empty.
    /// </summary>
    public static int? Poll()
    {
        if (!enabled)
        {
            return null;
        }

        while (GetNumberOfConsoleInputEvents(handle, out var count) && count > 0)
        {
            if (!PeekConsoleInput(handle, out var record, 1, out var peeked) || peeked == 0)
            {
                return null;
            }

            if (record.EventType == KeyEventType && IsRealKeyPress(record))
            {
                return 0;
            }

            ReadConsoleInput(handle, out _, 1, out _);
            if (record.EventType == MouseEventType && (record.Word4 & MouseWheeled) != 0)
            {
                return (int)record.Word2 > 0 ? -1 : 1;
            }
        }

        return null;
    }

    private static bool IsRealKeyPress(InputRecord record)
    {
        var keyDown = record.Word1 != 0;
        var virtualKey = record.Word2 >> 16;
        var character = record.Word3 >> 16;
        var modifierOnly = virtualKey is 16 or 17 or 18 or 20 or 91 or 92 or 144 or 145;
        return keyDown && (character != 0 || !modifierOnly);
    }
}
