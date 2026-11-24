param([long] $ExcelHwnd, [string] $Password)
$ErrorActionPreference = 'Stop'

if (-not ('VbaPasswordWindow' -as [type])) { Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class VbaPasswordWindow {
    private delegate bool EnumCallback(IntPtr window, IntPtr state);
    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumCallback callback, IntPtr state);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumCallback callback, IntPtr state);
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr window, int index);
    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
    private static extern IntPtr SetText(IntPtr window, uint message, IntPtr unused, string text);
    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wparam, IntPtr lparam);

    public static int ProcessId(long handle) {
        uint id;
        GetWindowThreadProcessId(new IntPtr(handle), out id);
        return (int)id;
    }
    public static long[][] Windows(int processId) {
        var result = new List<long[]>();
        EnumWindows((window, state) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner == (uint)processId && IsWindowVisible(window)) {
                var children = new List<IntPtr>();
                EnumChildWindows(window, (child, unused) => { children.Add(child); return true; }, IntPtr.Zero);
                IntPtr password = IntPtr.Zero;
                IntPtr confirm = IntPtr.Zero;
                foreach (var child in children) {
                    var name = new StringBuilder(128);
                    GetClassName(child, name, name.Capacity);
                    if (name.ToString().IndexOf("EDIT", StringComparison.OrdinalIgnoreCase) >= 0 &&
                        (GetWindowLong(child, -16) & 0x20) != 0) password = child;
                    if (name.ToString().IndexOf("BUTTON", StringComparison.OrdinalIgnoreCase) >= 0) {
                        var caption = new StringBuilder(128);
                        GetWindowText(child, caption, caption.Capacity);
                        if (GetDlgCtrlID(child) == 1 ||
                            string.Equals(caption.ToString(), "OK", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(caption.ToString(), "Aceptar", StringComparison.OrdinalIgnoreCase)) confirm = child;
                    }
                }
                if (password != IntPtr.Zero) result.Add(new long[] { window.ToInt64(), password.ToInt64(), confirm.ToInt64() });
            }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
    public static string Title(long handle) {
        var name = new StringBuilder(256);
        GetWindowText(new IntPtr(handle), name, name.Capacity);
        return name.ToString();
    }
    public static long[] PropertyDialogs(int processId, long excludedHandle) {
        var result = new List<long>();
        EnumWindows((window, state) => {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner == (uint)processId && IsWindowVisible(window) && window.ToInt64() != excludedHandle) {
                string title = Title(window.ToInt64());
                if (title.IndexOf("propert", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    title.IndexOf("propiedad", StringComparison.OrdinalIgnoreCase) >= 0)
                    result.Add(window.ToInt64());
            }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
    public static void Enter(long edit, long button, string password) {
        SetText(new IntPtr(edit), 0x000C, IntPtr.Zero, "");
        foreach (char character in password) {
            SendMessage(new IntPtr(edit), 0x0102, new IntPtr((int)character), IntPtr.Zero);
        }
        if (button == 0) throw new InvalidOperationException("No se encontro el boton Aceptar del cuadro VBA.");
        SendMessage(new IntPtr(button), 0x00F5, IntPtr.Zero, IntPtr.Zero);
    }
    public static void Close(long handle) {
        SendMessage(new IntPtr(handle), 0x0010, IntPtr.Zero, IntPtr.Zero);
    }
}
'@ }

$processId = [VbaPasswordWindow]::ProcessId($ExcelHwnd)
if (-not $processId) { throw 'No se identifico el proceso de Excel.' }
$deadline = [DateTime]::UtcNow.AddSeconds(30)
$dialog = $null
while ([DateTime]::UtcNow -lt $deadline -and -not $dialog) {
    $candidates = @([VbaPasswordWindow]::Windows($processId))
    if ($candidates.Count -eq 1) { $dialog = $candidates[0] }
    if (-not $dialog) { Start-Sleep -Milliseconds 150 }
}
if (-not $dialog) { throw 'No se encontro un unico cuadro de contrasena del proyecto VBA.' }
[VbaPasswordWindow]::Enter($dialog[1], $dialog[2], $Password)

$deadline = [DateTime]::UtcNow.AddSeconds(5)
while ([DateTime]::UtcNow -lt $deadline) {
    $properties = @([VbaPasswordWindow]::PropertyDialogs($processId, $dialog[0]))
    if ($properties.Count -eq 1) {
        [VbaPasswordWindow]::Close($properties[0])
        return 'Clave enviada; propiedades cerradas.'
    }
    Start-Sleep -Milliseconds 150
}
return 'Clave enviada.'
