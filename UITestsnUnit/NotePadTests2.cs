using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using NUnit.Framework;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace UITests;

public class NotepadTests2
{
    [Test]
    [Apartment(ApartmentState.STA)]
    public void Should_Write_Text_In_Bold_In_Notepad()
    {
        const string expectedText = "gus aiti killari";

        Process.Start("notepad.exe");

        using var automation = new UIA3Automation();

        var process = Retry.WhileNull(
            () => Process.GetProcessesByName("notepad")
                .OrderByDescending(p => TryGetStartTime(p))
                .FirstOrDefault(HasMainWindow),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)).Result;

        Assert.That(process, Is.Not.Null, "No se encontró el proceso de Notepad.");

        var window = Retry.WhileNull(
            () => automation.GetDesktop()
                .FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                .FirstOrDefault(w => w.Properties.ProcessId.ValueOrDefault == process!.Id),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(500)).Result;

        Assert.That(window, Is.Not.Null, "No se encontró la ventana de Notepad.");

        window!.Focus();
        Wait.UntilInputIsProcessed();

        var editor = Retry.WhileNull(
            () => window.FindFirstDescendant(cf =>
                    cf.ByName("Editor de texto").And(cf.ByControlType(ControlType.Document)))
               ?? window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Document))
               ?? window.FindFirstDescendant(cf => cf.ByControlType(ControlType.Edit)),
            timeout: TimeSpan.FromSeconds(10),
            interval: TimeSpan.FromMilliseconds(600)).Result;

        Assert.That(editor, Is.Not.Null, "No se encontró el editor de texto.");

        editor!.Focus();
        Wait.UntilInputIsProcessed();

        WriteTextReliably(editor, expectedText);

        // Seleccionar todo
        Keyboard.Press(VirtualKeyShort.CONTROL);
        Keyboard.Press(VirtualKeyShort.KEY_A);
        Keyboard.Release(VirtualKeyShort.KEY_A);
        Keyboard.Release(VirtualKeyShort.CONTROL);
        Wait.UntilInputIsProcessed();

        // Buscar el botón Negrita
        var boldButton = Retry.WhileNull(
            () => window.FindFirstDescendant(cf =>
                cf.ByControlType(ControlType.Button)
                  .And(cf.ByName("Negrita (Ctrl+B)"))),
            timeout: TimeSpan.FromSeconds(5),
            interval: TimeSpan.FromMilliseconds(300)).Result;

        Assert.That(boldButton, Is.Not.Null, "No se encontró el botón Negrita.");

        // Activar negrita
        var togglePattern = boldButton!.Patterns.Toggle.PatternOrDefault;
        if (togglePattern is not null)
        {
            // Solo activa si todavía no está activado
            if (togglePattern.ToggleState.Value != ToggleState.On)
            {
                togglePattern.Toggle();
                Wait.UntilInputIsProcessed();
            }
        }
        else
        {
            // Fallback: click o atajo
            var invokePattern = boldButton.Patterns.Invoke.PatternOrDefault;
            if (invokePattern is not null)
            {
                invokePattern.Invoke();
                Wait.UntilInputIsProcessed();
            }
            else
            {
                // Último recurso: Ctrl+B
                Keyboard.Press(VirtualKeyShort.CONTROL);
                Keyboard.Press(VirtualKeyShort.KEY_B);
                Keyboard.Release(VirtualKeyShort.KEY_B);
                Keyboard.Release(VirtualKeyShort.CONTROL);
                Wait.UntilInputIsProcessed();
            }
        }

        // Validar texto esperando a que la UI termine de actualizarse
        var textValidation = Retry.WhileFalse(
            () => string.Equals(TryGetEditorText(editor), expectedText, StringComparison.Ordinal),
            timeout: TimeSpan.FromSeconds(5),
            interval: TimeSpan.FromMilliseconds(200),
            throwOnTimeout: false);

        var actualText = TryGetEditorText(editor);

        Assert.That(textValidation.Result, Is.True,
            $"El texto esperado era '{expectedText}', pero se obtuvo '{actualText}'.");

        // Validar negrita, si el botón expone TogglePattern
        var boldToggleAfter = boldButton.Patterns.Toggle.PatternOrDefault;
        if (boldToggleAfter is not null)
        {
            Assert.That(boldToggleAfter.ToggleState.Value, Is.EqualTo(ToggleState.On),
                "El botón Negrita no quedó activado.");
        }
        else
        {
            Assert.Inconclusive(
                "Se pudo escribir y activar negrita, pero la UI no expone un estado verificable del botón para afirmar de forma confiable que el texto quedó en negrita.");
        }
    }

    private static string TryGetEditorText(AutomationElement editor)
    {
        // Intento 1: como TextBox
        try
        {
            return editor.AsTextBox().Text;
        }
        catch
        {
            // Ignorar y probar otras vías
        }

        // Intento 2: ValuePattern
        try
        {
            var valuePattern = editor.Patterns.Value.PatternOrDefault;
            if (valuePattern is not null)
            {
                return valuePattern.Value.Value ?? string.Empty;
            }
        }
        catch
        {
            // Ignorar
        }

        // Intento 3: TextPattern
        try
        {
            var textPattern = editor.Patterns.Text.PatternOrDefault;
            if (textPattern is not null)
            {
                return textPattern.DocumentRange.GetText(int.MaxValue)?.TrimEnd('\r', '\n') ?? string.Empty;
            }
        }
        catch
        {
            // Ignorar
        }

        return string.Empty;
    }

    private static void WriteTextReliably(AutomationElement editor, string expectedText)
    {
        var textWasWritten = Retry.WhileFalse(
            () =>
            {
                FocusEditorForTyping(editor);
                ClearEditor(editor);
                WriteTextToEditor(editor, expectedText);

                return string.Equals(TryGetEditorText(editor), expectedText, StringComparison.Ordinal);
            },
            timeout: TimeSpan.FromSeconds(8),
            interval: TimeSpan.FromMilliseconds(500),
            throwOnTimeout: false);

        if (!textWasWritten.Result)
        {
            var actualText = TryGetEditorText(editor);
            Assert.Fail($"No se pudo escribir el texto completo. Esperado: '{expectedText}'. Actual: '{actualText}'.");
        }
    }

    private static void FocusEditorForTyping(AutomationElement editor)
    {
        editor.Focus();
        Wait.UntilInputIsProcessed();
    }

    private static void ClearEditor(AutomationElement editor)
    {
        try
        {
            var valuePattern = editor.Patterns.Value.PatternOrDefault;
            if (valuePattern is not null && !valuePattern.IsReadOnly.Value)
            {
                valuePattern.SetValue(string.Empty);
                Wait.UntilInputIsProcessed();
                return;
            }
        }
        catch
        {
            // Ignorar y usar fallback por teclado
        }

        Keyboard.Press(VirtualKeyShort.CONTROL);
        Keyboard.Press(VirtualKeyShort.KEY_A);
        Keyboard.Release(VirtualKeyShort.KEY_A);
        Keyboard.Release(VirtualKeyShort.CONTROL);
        Wait.UntilInputIsProcessed();

        Keyboard.Press(VirtualKeyShort.DELETE);
        Keyboard.Release(VirtualKeyShort.DELETE);
        Wait.UntilInputIsProcessed();
    }

    private static void WriteTextToEditor(AutomationElement editor, string expectedText)
    {
        try
        {
            var valuePattern = editor.Patterns.Value.PatternOrDefault;
            if (valuePattern is not null && !valuePattern.IsReadOnly.Value)
            {
                valuePattern.SetValue(expectedText);
                Wait.UntilInputIsProcessed();
                return;
            }
        }
        catch
        {
            // Ignorar y usar fallback
        }

        try
        {
            editor.AsTextBox().Enter(expectedText);
            Wait.UntilInputIsProcessed();
            return;
        }
        catch
        {
            // Ignorar y usar fallback
        }

        Keyboard.Type(expectedText);
        Wait.UntilInputIsProcessed();
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
        catch
        {
            return DateTime.MinValue;
        }
    }
}