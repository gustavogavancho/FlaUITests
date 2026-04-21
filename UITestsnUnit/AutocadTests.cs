using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.Tools;
using FlaUI.UIA3;
using NUnit.Framework;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace UITests;

public class AutocadTests
{
    [Test]
    [Apartment(ApartmentState.STA)]
    public void Should_Draw_Circle_In_AutoCAD()
    {
        const double centerX = 10d;
        const double centerY = 15d;
        const double radius = 5d;

        using var _ = ComMessageFilter.Register();

        object? autoCadApp = null;
        object? document = null;
        object? modelSpace = null;
        object? circle = null;

        try
        {
            autoCadApp = CreateAutoCadApplication();
            dynamic app = autoCadApp;
            InvokeWithRetry(() => app.Visible = true);

            using var automation = new UIA3Automation();

            var process = WaitForAutoCadProcess();
            Assert.That(process, Is.Not.Null, "No se encontró el proceso de AutoCAD.");

            var window = WaitForMainWindow(automation, process!.Id);
            Assert.That(window, Is.Not.Null, "No se encontró la ventana principal de AutoCAD.");

            window!.Focus();
            Wait.UntilInputIsProcessed();

            document = WaitForDocument(app);
            dynamic activeDocument = document;

            modelSpace = InvokeWithRetry(() => (object)activeDocument.ModelSpace);
            dynamic drawingSpace = modelSpace;

            var initialEntityCount = InvokeWithRetry(() => Convert.ToInt32(drawingSpace.Count));

            circle = InvokeWithRetry(() => (object)drawingSpace.AddCircle(new double[] { centerX, centerY, 0d }, radius));
            dynamic createdCircle = circle;

            TryRegen(activeDocument);

            var finalEntityCount = InvokeWithRetry(() => Convert.ToInt32(drawingSpace.Count));
            Assert.That(finalEntityCount, Is.EqualTo(initialEntityCount + 1),
                "AutoCAD no agregó el círculo al ModelSpace.");

            Assert.That(InvokeWithRetry(() => (string)createdCircle.ObjectName), Is.EqualTo("AcDbCircle"),
                "La entidad creada no es un círculo.");
            Assert.That(InvokeWithRetry(() => Convert.ToDouble(createdCircle.Radius)), Is.EqualTo(radius).Within(0.001),
                "El radio del círculo no coincide con el esperado.");

            var center = ToDoubleArray(InvokeWithRetry(() => (object)createdCircle.Center));
            Assert.That(center.Length, Is.GreaterThanOrEqualTo(3),
                "No se pudo leer el centro del círculo creado.");
            Assert.That(center[0], Is.EqualTo(centerX).Within(0.001),
                "La coordenada X del centro no coincide.");
            Assert.That(center[1], Is.EqualTo(centerY).Within(0.001),
                "La coordenada Y del centro no coincide.");
            Assert.That(center[2], Is.EqualTo(0d).Within(0.001),
                "La coordenada Z del centro no coincide.");
        }
        finally
        {
            TryCloseDocument(document);
            TryQuitAutoCad(autoCadApp);

            ReleaseComObject(circle);
            ReleaseComObject(modelSpace);
            ReleaseComObject(document);
            ReleaseComObject(autoCadApp);
        }
    }

    private static object CreateAutoCadApplication()
    {
        var autoCadType = Type.GetTypeFromProgID("AutoCAD.Application");
        if (autoCadType is null)
        {
            Assert.Ignore("AutoCAD no está instalado o no expone automatización COM en este equipo.");
        }

        try
        {
            return Activator.CreateInstance(autoCadType!)
                ?? throw new InvalidOperationException("No se pudo crear la instancia de AutoCAD.");
        }
        catch (COMException ex)
        {
            Assert.Fail($"No se pudo iniciar AutoCAD mediante COM: {ex.Message}");
            throw;
        }
    }

    private static Process? WaitForAutoCadProcess()
    {
        return Retry.WhileNull(
            () => Process.GetProcessesByName("acad")
                .OrderByDescending(TryGetStartTime)
                .FirstOrDefault(HasMainWindow),
            timeout: TimeSpan.FromSeconds(60),
            interval: TimeSpan.FromSeconds(1)).Result;
    }

    private static AutomationElement? WaitForMainWindow(UIA3Automation automation, int processId)
    {
        return Retry.WhileNull(
            () => automation.GetDesktop()
                .FindAllChildren(cf => cf.ByControlType(ControlType.Window))
                .FirstOrDefault(window => window.Properties.ProcessId.ValueOrDefault == processId),
            timeout: TimeSpan.FromSeconds(60),
            interval: TimeSpan.FromSeconds(1)).Result;
    }

    private static object WaitForDocument(dynamic app)
    {
        var timeout = DateTime.UtcNow.AddSeconds(60);

        while (DateTime.UtcNow < timeout)
        {
            try
            {
                if (app.Documents.Count == 0)
                {
                    app.Documents.Add();
                }

                var document = app.ActiveDocument;
                if (document is not null)
                {
                    return document;
                }
            }
            catch (COMException)
            {
            }

            Thread.Sleep(1000);
        }

        Assert.Fail("AutoCAD no abrió un dibujo listo para automatizar dentro del tiempo esperado.");
        throw new InvalidOperationException();
    }

    private static void TryRegen(dynamic document)
    {
        try
        {
            InvokeWithRetry(() => document.Regen(1));
        }
        catch (COMException)
        {
        }
    }

    private static void InvokeWithRetry(Action action)
    {
        InvokeWithRetry(() =>
        {
            action();
            return true;
        });
    }

    private static T InvokeWithRetry<T>(Func<T> action)
    {
        var timeout = DateTime.UtcNow.AddSeconds(30);
        COMException? lastException = null;

        while (DateTime.UtcNow < timeout)
        {
            try
            {
                return action();
            }
            catch (COMException ex) when (IsCallRejected(ex))
            {
                lastException = ex;
                Thread.Sleep(250);
            }
        }

        if (lastException is not null)
        {
            throw lastException;
        }

        throw new TimeoutException("La operación COM excedió el tiempo de espera.");
    }

    private static bool IsCallRejected(COMException exception)
    {
        return exception.HResult == unchecked((int)0x80010001);
    }

    private static double[] ToDoubleArray(object values)
    {
        if (values is not Array array)
        {
            return Array.Empty<double>();
        }

        return array.Cast<object>()
            .Select(Convert.ToDouble)
            .ToArray();
    }

    private static void TryCloseDocument(object? document)
    {
        if (document is null)
        {
            return;
        }

        try
        {
            ((dynamic)document).Close(false);
        }
        catch (COMException)
        {
        }
    }

    private static void TryQuitAutoCad(object? autoCadApp)
    {
        if (autoCadApp is null)
        {
            return;
        }

        try
        {
            ((dynamic)autoCadApp).Quit();
        }
        catch (COMException)
        {
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.FinalReleaseComObject(value);
        }
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

    private sealed class ComMessageFilter : IDisposable
    {
        private readonly IOleMessageFilter? _previousFilter;

        private ComMessageFilter(IOleMessageFilter? previousFilter)
        {
            _previousFilter = previousFilter;
        }

        public static ComMessageFilter Register()
        {
            var newFilter = new MessageFilterImpl();
            CoRegisterMessageFilter(newFilter, out var previousFilter);
            return new ComMessageFilter(previousFilter);
        }

        public void Dispose()
        {
            CoRegisterMessageFilter(_previousFilter, out _);
        }

        [DllImport("Ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

        [ComImport]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }

        private sealed class MessageFilterImpl : IOleMessageFilter
        {
            public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            {
                return 0;
            }

            public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            {
                return dwRejectType == 2 ? 100 : -1;
            }

            public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            {
                return 2;
            }
        }
    }
}
