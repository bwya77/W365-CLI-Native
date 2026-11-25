using System.Text;
using System.Text.RegularExpressions;
using Spectre.Console;

namespace W365Cli;

/// <summary>
/// Sits between Spectre/Console and the real terminal. Every screen redraws with
/// AnsiConsole.Clear() followed by many small writes; on a terminal shorter than the frame this
/// erases the screen, scrolls, and flashes (the top bar appears/disappears while arrowing).
/// This writer buffers output and presents each redraw as ONE write: cursor-home, the new frame
/// with erase-to-end-of-line per row, erase-below. No full-screen clear means no flash, and when
/// the frame is taller than the window it is clipped to the window (top bar pinned, selected
/// row kept in view) so the terminal never scrolls. Output that is not a full redraw (spinners,
/// typed characters) passes straight through unchanged.
/// </summary>
internal sealed partial class FrameWriter : TextWriter
{
    private const string ClearScreen = "\x1b[2J";
    private const int IdleFlushMs = 60;

    private readonly TextWriter _inner;
    private readonly StringBuilder _buffer = new();
    private readonly object _gate = new();
    private readonly Timer _idleTimer;
    private int _offset;

    private static FrameWriter? instance;

    private FrameWriter(TextWriter inner)
    {
        _inner = inner;
        _idleTimer = new Timer(_ => Present(force: false), null, Timeout.Infinite, Timeout.Infinite);
    }

    public override Encoding Encoding => _inner.Encoding;

    /// <summary>
    /// Routes Console and AnsiConsole through a FrameWriter. No-op when not attached to an
    /// interactive ANSI terminal (redirected/piped output keeps the default behavior).
    /// </summary>
    public static void Install()
    {
        try
        {
            var profile = AnsiConsole.Profile;
            if (instance is not null || Console.IsOutputRedirected || Console.IsInputRedirected ||
                !profile.Capabilities.Ansi || !profile.Capabilities.Interactive)
            {
                return;
            }

            var writer = new FrameWriter(Console.Out);
            Console.SetOut(writer);
            AnsiConsole.Console = AnsiConsole.Create(new AnsiConsoleSettings
            {
                Ansi = AnsiSupport.Yes,
                ColorSystem = profile.Capabilities.ColorSystem switch
                {
                    ColorSystem.NoColors => ColorSystemSupport.NoColors,
                    ColorSystem.Legacy => ColorSystemSupport.Legacy,
                    ColorSystem.Standard => ColorSystemSupport.Standard,
                    ColorSystem.EightBit => ColorSystemSupport.EightBit,
                    _ => ColorSystemSupport.TrueColor
                },
                Interactive = InteractionSupport.Yes,
                Out = new LiveSizeOutput(writer)
            });
            instance = writer;

            // Alternate screen buffer: the app owns the whole viewport and leaves no scrollback
            // of old frames behind (scrolling up used to reveal stale, garbled top bars).
            writer._inner.Write("\x1b[?1049h");
            writer._inner.Flush();
            ConsoleMouse.Enable();

            void Restore()
            {
                ConsoleMouse.Disable();
                instance?.PresentDiscard();
                writer._inner.Write("\x1b[?1049l");
                writer._inner.Flush();
            }

            AppDomain.CurrentDomain.ProcessExit += (_, _) => Restore();
            AppDomain.CurrentDomain.UnhandledException += (_, _) => Restore();
        }
        catch
        {
            // Terminal doesn't support what we need — keep the default console.
        }
    }

    /// <summary>Flushes any buffered frame now; call right before blocking on a key press.</summary>
    public static void PresentPending() => instance?.Present();

    public override void Write(char value) => Append(value.ToString());

    public override void Write(string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            Append(value);
        }
    }

    public override void Write(char[] buffer, int index, int count) => Append(new string(buffer, index, count));

    public override void Write(ReadOnlySpan<char> buffer) => Append(buffer.ToString());

    public override void Flush() => Present();

    private void Append(string text)
    {
        lock (_gate)
        {
            _buffer.Append(text);
            _idleTimer.Change(IdleFlushMs, Timeout.Infinite);
        }
    }

    /// <summary>Drops buffered output (used at shutdown, where a final redraw is pointless).</summary>
    private void PresentDiscard()
    {
        lock (_gate)
        {
            _buffer.Clear();
        }
    }

    private void Present(bool force = true)
    {
        lock (_gate)
        {
            if (_buffer.Length == 0)
            {
                return;
            }

            var text = _buffer.ToString();
            var clearAt = text.LastIndexOf(ClearScreen, StringComparison.Ordinal);
            string output;
            if (clearAt < 0)
            {
                output = text;
            }
            else
            {
                var frame = LeadingClearRemainder().Replace(text[(clearAt + ClearScreen.Length)..], string.Empty);
                if (!force && string.IsNullOrWhiteSpace(frame))
                {
                    // The idle timer fired between Clear() and the first real content (a slow
                    // step mid-render). Presenting now would show a blank screen, so hold the
                    // buffer until content arrives or the next key read forces it out.
                    return;
                }

                if (CursorMovement().IsMatch(frame))
                {
                    // Contains its own cursor control (live widgets) — don't try to compose it.
                    output = text[clearAt..];
                }
                else
                {
                    output = Compose(frame, SafeWindowHeight(), ref _offset);
                }
            }

            _buffer.Clear();
            _inner.Write(output);
            _inner.Flush();
        }
    }

    /// <summary>
    /// Tracks window size changes. Returns true when the size changed since the last call.
    /// </summary>
    public static bool SyncSize(ref (int Width, int Height) lastSize)
    {
        var size = (SafeWindowWidth(), SafeWindowHeight());
        var changed = lastSize != size && lastSize != default;
        lastSize = size;
        return changed;
    }

    private static int SafeWindowWidth()
    {
        try
        {
            return Console.WindowWidth;
        }
        catch
        {
            return 0;
        }
    }

    private static int SafeWindowHeight()
    {
        try
        {
            return Console.WindowHeight > 0 ? Console.WindowHeight : int.MaxValue;
        }
        catch
        {
            return int.MaxValue;
        }
    }

    /// <summary>
    /// Builds the single-write output for one full redraw. <paramref name="frame"/> is the content
    /// emitted after Clear(). <paramref name="offset"/> is the sticky scroll position used only
    /// when the frame is taller than <paramref name="height"/>.
    /// </summary>
    internal static string Compose(string frame, int height, ref int offset)
    {
        var trailingNewline = frame.EndsWith('\n');
        var content = trailingNewline ? frame[..^1] : frame;
        var lines = content.Split('\n').Select(line => line.TrimEnd('\r')).ToArray();

        IEnumerable<string> visible = lines;
        var keepTrailingNewline = trailingNewline;
        if (lines.Length + (trailingNewline ? 1 : 0) > height)
        {
            var pinned = lines.Length > 0 && Background().IsMatch(lines[0]);
            var bodyStart = pinned ? 1 : 0;
            var windowSize = Math.Max(1, height - bodyStart);
            var bodyCount = lines.Length - bodyStart;

            var selected = -1;
            for (var index = bodyStart; index < lines.Length; index++)
            {
                if (Background().IsMatch(lines[index]))
                {
                    selected = index - bodyStart;
                    break;
                }
            }

            if (selected < 0)
            {
                offset = 0;
            }
            else
            {
                var margin = Math.Min(2, (windowSize - 1) / 2);
                if (selected <= windowSize - 1 - margin)
                {
                    // Selection is reachable without scrolling: show the screen from the top
                    // (title, summary, table header) instead of a stale scrolled position.
                    offset = 0;
                }
                else if (selected < offset + margin)
                {
                    offset = selected - margin;
                }
                else if (selected > offset + windowSize - 1 - margin)
                {
                    offset = selected - (windowSize - 1 - margin);
                }
            }

            offset = Math.Clamp(offset, 0, Math.Max(0, bodyCount - windowSize));
            var body = lines.Skip(bodyStart + offset).Take(windowSize);
            visible = pinned ? lines.Take(1).Concat(body) : body;
            keepTrailingNewline = false;
        }
        else
        {
            offset = 0;
        }

        var builder = new StringBuilder("\x1b[H");
        builder.AppendJoin("\x1b[K\n", visible).Append("\x1b[K");
        if (keepTrailingNewline)
        {
            builder.Append('\n');
        }

        return builder.Append("\x1b[J").ToString();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            Present();
            _idleTimer.Dispose();
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// Spectre output that reports the terminal's current size on every query. The stock
    /// AnsiConsoleOutput only does that for the real stdout writer; wrapped in a FrameWriter it
    /// would fall back to a fixed 80x24 and clip wide screens (e.g. Cloud PC details) no matter
    /// how large the window is.
    /// </summary>
    private sealed class LiveSizeOutput(TextWriter writer) : IAnsiConsoleOutput
    {
        public TextWriter Writer => writer;

        public bool IsTerminal => true;

        // One column short of the window: drawing into the last column makes terminals wrap or
        // clip the right border (what the stock Spectre output also avoids).
        public int Width => SafeWindowWidth() is var width and > 1 ? width - 1 : 80;

        public int Height => SafeWindowHeight() is var height and > 0 and < int.MaxValue ? height : 24;

        public void SetEncoding(Encoding encoding)
        {
            if (Console.OutputEncoding.CodePage != encoding.CodePage)
            {
                Console.OutputEncoding = encoding;
            }
        }
    }

    [GeneratedRegex(@"^(\x1b\[3J)?(\x1b\[1;1H)?")]
    private static partial Regex LeadingClearRemainder();

    // Any CSI cursor/erase sequence other than colors (SGR 'm') and cursor show/hide.
    [GeneratedRegex(@"\x1b\[[0-9;]*[ABCDEFGHJKSTfsu]")]
    private static partial Regex CursorMovement();

    // SGR containing an extended background color (the selected-row / nav-bar highlight).
    [GeneratedRegex(@"\x1b\[(?:[0-9;]*;)?48;[0-9;]*m")]
    private static partial Regex Background();
}
