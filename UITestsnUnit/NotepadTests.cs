using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Tools;
using FlaUI.UIA3;
using System.ComponentModel;
using NUnit.Framework;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace UITests;

public class NotepadTests
{
    [Test]
    [Apartment(ApartmentState.STA)]
    public void Should_Write_Text_In_Notepad()
    {
        Process.Start("notepad.exe");

        using var automation = new UIA3Automation();

        var process = Retry.WhileNull(
            () => Process.GetProcessesByName("notepad")
                .OrderByDescending(p => TryGetStartTime(p))
                .FirstOrDefault(HasMainWindow),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)).Result;

        Assert.That(process, Is.Not.Null);

        var window = Retry.WhileNull(
            () => automation.GetDesktop()
                .FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                .FirstOrDefault(w => w.Properties.ProcessId.ValueOrDefault == process!.Id),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)).Result;

        Assert.That(window, Is.Not.Null);

        window!.Focus();

        var editor = Retry.WhileNull(
            () => window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Document))
               ?? window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit)),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)).Result;

        Assert.That(editor, Is.Not.Null);

        editor!.AsTextBox().Enter("Hola desde FlaUI");
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