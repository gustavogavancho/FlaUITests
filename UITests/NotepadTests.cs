using FlaUI.Core.Tools;
using FlaUI.UIA3;

namespace UITests;

public class NotepadTests
{
    [Test]
    public void Should_Write_Text_In_Notepad()
    {
        using var app = Application.Launch("notepad.exe");
        using var automation = new UIA3Automation();

        var window = Retry.WhileNull(
            () => app.GetMainWindow(automation),
            timeout: TimeSpan.FromSeconds(5)).Result;

        Assert.That(window, Is.Not.Null);

        var editor = window.FindFirstDescendant(cf =>
            cf.ByControlType(FlaUI.Core.Definitions.ControlType.Document));

        Assert.That(editor, Is.Not.Null);

        editor.AsTextBox().Enter("Hola desde FlaUI");
    }
}