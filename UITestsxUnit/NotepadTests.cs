using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace UITests;

public class NotepadTests
{
    [StaFact]
    public void Should_Write_Text_In_Notepad()
    {
        Process.Start("notepad.exe");
        using var automation = new UIA3Automation();

        var process = Retry.WhileNull(
            () => Process.GetProcessesByName("notepad")
                .OrderByDescending(p => TryGetStartTime(p))
                .FirstOrDefault(HasMainWindow),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)
        ).Result;

        Assert.NotNull(process);

        var window = Retry.WhileNull(
            () => automation.GetDesktop()
                .FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                .FirstOrDefault(w => w.Properties.ProcessId.ValueOrDefault == process!.Id),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)
        ).Result;

        Assert.NotNull(window);

        window.Focus();

        var editor = Retry.WhileNull(
            () => window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Document))
               ?? window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit)),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)
        ).Result;

        Assert.NotNull(editor);

        editor.AsTextBox().Enter("Hola desde FlaUI");
    }

    private static bool HasMainWindow(Process process)
    {
        try
        {
            process.Refresh();
            return !process.HasExited && process.MainWindowHandle != IntPtr.Zero;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (Win32Exception)
        {
            return false;
        }
    }

    private static DateTime TryGetStartTime(Process process)
    {
        try
        {
            return process.StartTime;
        }
        catch (Exception)
        {
            return DateTime.MinValue;
        }
    }
}