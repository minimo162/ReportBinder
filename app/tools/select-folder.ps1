# ASCII-only helper. Shows a modern Explorer-style folder picker without a visible console window.
$ErrorActionPreference = 'Stop'

function Decode-EnvValue([string]$Name, [string]$Fallback) {
    $value = [Environment]::GetEnvironmentVariable($Name, 'Process')
    if ([string]::IsNullOrWhiteSpace($value)) { return $Fallback }
    try { return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($value)) }
    catch { return $Fallback }
}

$Title = Decode-EnvValue 'REPORTBINDER_PICKER_TITLE_B64' 'Select folder'
$InitialDir = Decode-EnvValue 'REPORTBINDER_PICKER_INITIAL_B64' ([Environment]::GetFolderPath('MyDocuments'))
$OutputPath = Decode-EnvValue 'REPORTBINDER_PICKER_OUTPUT_B64' ''

function Write-SelectedPath([string]$Path) {
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        Set-Content -LiteralPath $OutputPath -Value ([string]$Path) -Encoding UTF8
    } else {
        Write-Output ([string]$Path)
    }
}

$owner = $null
try {
    if ([string]::IsNullOrWhiteSpace($InitialDir) -or -not (Test-Path -LiteralPath $InitialDir)) {
        $InitialDir = [Environment]::GetFolderPath('MyDocuments')
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Invisible owner window. This helps the dialog appear in front of the browser
    # while keeping the PowerShell console hidden.
    $owner = New-Object System.Windows.Forms.Form
    $owner.Text = 'ReportBinder'
    $owner.ShowInTaskbar = $false
    $owner.TopMost = $true
    $owner.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    try { $owner.Opacity = 0 } catch { }
    $owner.Show()
    $owner.Activate()

    # Use the modern Explorer-style folder picker.
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

    $source = @"
using System;
using System.Runtime.InteropServices;

namespace ReportBinder
{
    public static class NativeFolderPicker
    {
        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint FOS_NOCHANGEDIR = 0x00000008;
        private const uint FOS_PATHMUSTEXIST = 0x00000800;
        private const int ERROR_CANCELLED = unchecked((int)0x800704C7);

        [ComImport]
        [Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        private class FileOpenDialogRCW { }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
        private interface IFileOpenDialog
        {
            [PreserveSig]
            int Show(IntPtr parent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint fos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppsai);
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(SIGDN sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        private enum SIGDN : uint
        {
            FILESYSPATH = 0x80058000
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void SHCreateItemFromParsingName(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            IntPtr pbc,
            ref Guid riid,
            out IShellItem ppv);

        public static string Pick(string title, string initialDirectory, IntPtr ownerHandle)
        {
            IFileOpenDialog dialog = null;
            IShellItem initialItem = null;
            IShellItem resultItem = null;
            try
            {
                dialog = (IFileOpenDialog)new FileOpenDialogRCW();

                uint options;
                dialog.GetOptions(out options);
                options = options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_NOCHANGEDIR;
                dialog.SetOptions(options);

                if (!String.IsNullOrWhiteSpace(title))
                {
                    dialog.SetTitle(title);
                }
                // Keep the default localized Explorer button label.

                if (!String.IsNullOrWhiteSpace(initialDirectory) && System.IO.Directory.Exists(initialDirectory))
                {
                    try
                    {
                        Guid shellItemGuid = typeof(IShellItem).GUID;
                        SHCreateItemFromParsingName(initialDirectory, IntPtr.Zero, ref shellItemGuid, out initialItem);
                        dialog.SetDefaultFolder(initialItem);
                        dialog.SetFolder(initialItem);
                    }
                    catch
                    {
                        // Ignore invalid or inaccessible initial paths.
                    }
                }

                int hr = dialog.Show(ownerHandle);
                if (hr == ERROR_CANCELLED)
                {
                    return String.Empty;
                }
                if (hr != 0)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                dialog.GetResult(out resultItem);
                IntPtr pszString = IntPtr.Zero;
                try
                {
                    resultItem.GetDisplayName(SIGDN.FILESYSPATH, out pszString);
                    return Marshal.PtrToStringUni(pszString) ?? String.Empty;
                }
                finally
                {
                    if (pszString != IntPtr.Zero)
                    {
                        Marshal.FreeCoTaskMem(pszString);
                    }
                }
            }
            finally
            {
                if (resultItem != null) Marshal.ReleaseComObject(resultItem);
                if (initialItem != null) Marshal.ReleaseComObject(initialItem);
                if (dialog != null) Marshal.ReleaseComObject(dialog);
            }
        }
    }
}
"@

    Add-Type -TypeDefinition $source -Language CSharp -ErrorAction Stop
    $selected = [ReportBinder.NativeFolderPicker]::Pick($Title, $InitialDir, $owner.Handle)
    if ($null -eq $selected) { $selected = '' }
    Write-SelectedPath $selected
    exit 0
} catch {
    try { Write-SelectedPath '' } catch { }
    exit 1
} finally {
    if ($owner) { try { $owner.Close(); $owner.Dispose() } catch { } }
}
