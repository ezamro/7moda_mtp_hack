using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using STATSTG = System.Runtime.InteropServices.ComTypes.STATSTG;

namespace MtpTool;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("7moda MTP Tool v1.0");
        Console.WriteLine("====================");

        if (args.Length == 0) { PrintHelp(); return; }

        try
        {
            var cmd = args[0].ToLower();
            switch (cmd)
            {
                case "list": case "detect": case "ls": ListDevices(); break;
                case "info":
                    if (args.Length < 2) { Console.WriteLine("Usage: 7moda info <index|name>"); return; }
                    ShowDeviceInfo(args[1]); break;
                case "dir": case "files":
                    if (args.Length < 2) { Console.WriteLine("Usage: 7moda dir <index|name> [path]"); return; }
                    ListFiles(args[1], args.Length > 2 ? args[2] : ""); break;
                case "get": case "pull":
                    if (args.Length < 4) { Console.WriteLine("Usage: 7moda get <index|name> <remote_path> <local_path>"); return; }
                    CopyFile(args[1], args[2], args[3]); break;
                case "tree":
                    if (args.Length < 2) { Console.WriteLine("Usage: 7moda tree <index|name> [path]"); return; }
                    TreeFiles(args[1], args.Length > 2 ? args[2] : "", 0); break;
                default: PrintHelp(); break;
            }
        }
        catch (Exception ex) { Console.WriteLine($"Error: {ex.Message}"); }
    }

    static void PrintHelp()
    {
        Console.WriteLine(@"
USAGE:
  7moda list                    - List connected MTP devices
  7moda info <index|name>       - Show detailed device info
  7moda dir <index|name> [path] - List files on device
  7moda tree <index|name> [p]   - Show full folder tree
  7moda get <idx> <src> <dst>   - Copy file from device

EXAMPLES:
  7moda list
  7moda info 0
  7moda dir 0 ""DCIM/Camera""
  7moda get 0 ""DCIM/photo.jpg"" C:\temp\photo.jpg
");
    }

    struct DeviceInfo
    {
        public int Index;
        public string Name;
        public string Type;
        public string DeviceType;
        public string TotalSize;
        public string FreeSpace;
        public dynamic Item;
        public dynamic Computer;
    }

    static dynamic GetShell()
    {
        var t = Type.GetTypeFromProgID("Shell.Application");
        return Activator.CreateInstance(t);
    }

    static List<DeviceInfo> GetMtpDevices()
    {
        var list = new List<DeviceInfo>();
        dynamic shell = null;
        dynamic computer = null;
        try
        {
            shell = GetShell();
            computer = shell.NameSpace(17);
            var items = computer.Items();
            int count = items.Count;
            int idx = 0;
            for (int i = 0; i < count; i++)
            {
                dynamic item = items.Item(i);
                if (item == null) continue;
                try
                {
                    string name = item.Name ?? "";
                    string type = "";
                    try { type = computer.GetDetailsOf(item, 2) ?? ""; } catch { }
                    if (string.IsNullOrEmpty(type))
                        try { type = item.ExtendedProperty("System.ItemTypeText") ?? ""; } catch { }
                    bool isFolder = false, isFileSystem = true;
                    try { isFolder = item.IsFolder; } catch { }
                    try { isFileSystem = item.IsFileSystem; } catch { }
                    bool isMtp = !isFileSystem && isFolder;
                    if (!isMtp && !string.IsNullOrEmpty(type) &&
                        (type.IndexOf("portable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         type.IndexOf("mtp", StringComparison.OrdinalIgnoreCase) >= 0))
                        isMtp = true;
                    if (!isMtp && !string.IsNullOrEmpty(name) &&
                        (name.IndexOf("phone", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("android", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("oppo", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("samsung", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("xiaomi", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("huawei", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("nokia", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("sony", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("lg ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("motorola", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("oneplus", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("google", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("pixel", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("tablet", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         name.IndexOf("camera", StringComparison.OrdinalIgnoreCase) >= 0))
                        isMtp = true;
                    var di = new DeviceInfo
                    {
                        Index = idx, Name = name, Type = type,
                        DeviceType = isMtp ? "MTP" : "Unknown",
                        Item = item, Computer = computer,
                        TotalSize = "", FreeSpace = ""
                    };
                    try { di.TotalSize = computer.GetDetailsOf(item, 11) ?? ""; } catch { }
                    try { di.FreeSpace = computer.GetDetailsOf(item, 6) ?? ""; } catch { }
                    if (isMtp) list.Add(di);
                    idx++;
                }
                catch { }
            }
        }
        finally
        {
            if (computer != null) try { Marshal.FinalReleaseComObject(computer); } catch { }
            if (shell != null) try { Marshal.FinalReleaseComObject(shell); } catch { }
        }
        return list;
    }

    static DeviceInfo? FindDevice(string selector)
    {
        var devices = GetMtpDevices();
        if (devices.Count == 0) return null;
        if (int.TryParse(selector, out int idx))
        {
            foreach (var d in devices)
                if (d.Index == idx) return d;
            Console.WriteLine($"Device index {idx} not found");
            return null;
        }
        foreach (var d in devices)
            if (d.Name.IndexOf(selector, StringComparison.OrdinalIgnoreCase) >= 0)
                return d;
        Console.WriteLine($"No device matching '{selector}' found");
        return null;
    }

    static void ListDevices()
    {
        var devices = GetMtpDevices();
        if (devices.Count == 0)
        {
            Console.WriteLine("No MTP devices detected.");
            Console.WriteLine("Make sure your device is connected and unlocked.");
            return;
        }
        Console.WriteLine($"Found {devices.Count} device(s):\n");
        foreach (var d in devices)
        {
            Console.WriteLine($"  [{d.Index}] {d.Name}");
            Console.WriteLine($"       Type: {d.Type}");
            if (!string.IsNullOrEmpty(d.TotalSize))
                Console.WriteLine($"       Size: {d.TotalSize}");
            if (!string.IsNullOrEmpty(d.FreeSpace))
                Console.WriteLine($"       Free: {d.FreeSpace}");
            Console.WriteLine();
        }
    }

    static void QueryWmiDeviceInfo(string deviceName)
    {
        try
        {
            string q = "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%" +
                       deviceName.Replace("'", "''") + "%'";
            foreach (var obj in new ManagementObjectSearcher(q).Get())
            {
                try
                {
                    string pnpId = obj["PNPDeviceID"]?.ToString() ?? "";
                    string status = obj["Status"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(pnpId))
                    {
                        Console.WriteLine($"  Device ID    : {pnpId}");
                        Console.WriteLine($"  Status       : {status}");
                    }
                    if (pnpId.Contains("VID_"))
                    {
                        int vStart = pnpId.IndexOf("VID_");
                        string vid = pnpId.Substring(vStart, 8);
                        Console.WriteLine($"  USB Vendor   : {vid}");
                    }
                }
                finally { obj.Dispose(); }
            }
        }
        catch { }
    }

    static void ShowDeviceInfo(string selector)
    {
        var device = FindDevice(selector);
        if (device == null) return;
        var d = device.Value;
        Console.WriteLine($"\nDevice: {d.Name}");
        Console.WriteLine(new string('=', 40));
        Console.WriteLine($"  Index      : {d.Index}");
        Console.WriteLine($"  Name       : {d.Name}");
        Console.WriteLine($"  Type       : {d.Type}");
        Console.WriteLine($"  Total Size : {d.TotalSize}");
        Console.WriteLine($"  Free Space : {d.FreeSpace}");
        try
        {
            dynamic shell = null, computer = null, deviceItem = null;
            try
            {
                shell = GetShell();
                computer = shell.NameSpace(17);
                deviceItem = computer.Items().Item(d.Index);
                if (deviceItem != null)
                {
                    string[] props = {
                        "System.ItemNameDisplay", "System.ItemTypeText",
                        "System.Device.ModelName", "System.Device.Manufacturer",
                        "System.Device.FriendlyName", "System.Device.SerialNumber",
                        "System.Device.BatteryLevel", "System.Devices.BatteryPercentage",
                        "System.Capacity", "System.FreeSpace", "System.FirmwareVersion"
                    };
                    bool hasExtra = false;
                    foreach (var prop in props)
                    {
                        try
                        {
                            object val = deviceItem.ExtendedProperty(prop);
                            if (val != null)
                            {
                                string str = val.ToString() ?? "";
                                if (!string.IsNullOrWhiteSpace(str))
                                {
                                    if (!hasExtra) { Console.WriteLine($"\n  --- Properties ---"); hasExtra = true; }
                                    string label = prop.Replace("System.", "").Replace("Device.", "").Replace("Devices.", "");
                                    Console.WriteLine($"  {label,-22}: {str}");
                                }
                            }
                        }
                        catch { }
                    }
                    if (d.DeviceType == "MTP")
                    {
                        try
                        {
                            dynamic folder = deviceItem.GetFolder;
                            if (folder != null)
                            {
                                int itemCount = folder.Items().Count;
                                Console.WriteLine($"\n  Root Items  : {itemCount}");
                                Marshal.FinalReleaseComObject(folder);
                            }
                        }
                        catch { }
                    }
                }
            }
            finally
            {
                if (deviceItem != null) try { Marshal.FinalReleaseComObject(deviceItem); } catch { }
                if (computer != null) try { Marshal.FinalReleaseComObject(computer); } catch { }
                if (shell != null) try { Marshal.FinalReleaseComObject(shell); } catch { }
            }
        }
        catch { }
        try { QueryWmiDeviceInfo(d.Name); } catch { }
        Console.WriteLine();
    }

    static void ListFiles(string selector, string path)
    {
        var device = FindDevice(selector);
        if (device == null) return;
        try
        {
            var d = device.Value;
            dynamic shell = GetShell();
            dynamic computer = shell.NameSpace(17);
            try
            {
                dynamic deviceItem = computer.Items().Item(d.Index);
                if (deviceItem == null) { Console.WriteLine("Cannot access device"); return; }
                dynamic deviceFolder = deviceItem.GetFolder;
                if (deviceFolder == null) { Console.WriteLine("Cannot open device folder"); return; }
                dynamic targetFolder = deviceFolder;
                if (!string.IsNullOrEmpty(path))
                {
                    targetFolder = NavigateToFolder(deviceFolder, path);
                    if (targetFolder == null) { Console.WriteLine($"Path '{path}' not found on device"); return; }
                }
                var items = targetFolder.Items();
                int count = items.Count;
                Console.WriteLine($"\n{d.Name}:{NormalizePath(path)}");
                Console.WriteLine(new string('-', 60));
                if (count == 0) Console.WriteLine("  (empty)");
                int fileCount = 0, dirCount = 0;
                for (int i = 0; i < count; i++)
                {
                    dynamic item = items.Item(i);
                    if (item == null) continue;
                    try
                    {
                        string name = item.Name ?? "";
                        bool isFolder = false;
                        try { isFolder = item.IsFolder; } catch { }
                        string date = "";
                        try { date = targetFolder.GetDetailsOf(item, 3) ?? ""; } catch { }
                        string size = "";
                        try { long sz = 0; try { sz = (long)item.ExtendedProperty("System.Size"); } catch { } if (sz > 0) size = FormatSize(sz); } catch { }
                        if (isFolder) dirCount++; else fileCount++;
                        Console.WriteLine($"  {(isFolder ? "[DIR]" : "     ")} {name,-35} {size,10}  {date}");
                    }
                    catch { }
                }
                Console.WriteLine(new string('-', 60));
                Console.WriteLine($"  {dirCount} dir(s), {fileCount} file(s)\n");
            }
            finally
            {
                if (computer != null) try { Marshal.FinalReleaseComObject(computer); } catch { }
                if (shell != null) try { Marshal.FinalReleaseComObject(shell); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"Error accessing device: {ex.Message}"); }
    }

    static void TreeFiles(string selector, string path, int depth)
    {
        if (depth > 4) return;
        var device = FindDevice(selector);
        if (device == null) return;
        try
        {
            var d = device.Value;
            dynamic shell = GetShell();
            dynamic computer = shell.NameSpace(17);
            try
            {
                dynamic deviceItem = computer.Items().Item(d.Index);
                if (deviceItem == null) return;
                dynamic deviceFolder = deviceItem.GetFolder;
                if (deviceFolder == null) return;
                dynamic targetFolder = deviceFolder;
                if (!string.IsNullOrEmpty(path))
                {
                    targetFolder = NavigateToFolder(deviceFolder, path);
                    if (targetFolder == null) return;
                }
                PrintTree(targetFolder, path, depth, selector);
                Marshal.FinalReleaseComObject(targetFolder);
                Marshal.FinalReleaseComObject(deviceFolder);
                Marshal.FinalReleaseComObject(deviceItem);
            }
            finally
            {
                if (computer != null) try { Marshal.FinalReleaseComObject(computer); } catch { }
                if (shell != null) try { Marshal.FinalReleaseComObject(shell); } catch { }
            }
        }
        catch { }
    }

    static void PrintTree(dynamic folder, string currentPath, int depth, string selector)
    {
        if (depth > 4) return;
        var items = folder.Items();
        for (int i = 0; i < items.Count; i++)
        {
            dynamic item = items.Item(i);
            if (item == null) continue;
            try
            {
                string name = item.Name ?? "";
                bool isFolder = false;
                try { isFolder = item.IsFolder; } catch { }
                string indent = new string(' ', depth * 2);
                if (isFolder)
                {
                    Console.WriteLine($"{indent}[{name}]");
                    try
                    {
                        dynamic subFolder = item.GetFolder;
                        if (subFolder != null)
                        {
                            string subPath = string.IsNullOrEmpty(currentPath) ? name : $"{currentPath}/{name}";
                            PrintTree(subFolder, subPath, depth + 1, selector);
                            Marshal.FinalReleaseComObject(subFolder);
                        }
                    }
                    catch { }
                }
                else Console.WriteLine($"{indent}  {name}");
            }
            catch { }
        }
    }

    static dynamic FindItemInFolder(dynamic folder, string name)
    {
        try
        {
            var items = folder.Items();
            for (int i = 0; i < items.Count; i++)
            {
                dynamic item = items.Item(i);
                if (item == null) continue;
                try
                {
                    string n = item.Name ?? "";
                    if (n.Trim().Equals(name, StringComparison.OrdinalIgnoreCase) ||
                        n.Trim().Equals(name.Trim(), StringComparison.OrdinalIgnoreCase))
                        return item;
                }
                catch { }
                try { Marshal.ReleaseComObject(item); } catch { }
            }
        }
        catch { }
        return null;
    }

    static dynamic NavigateToFolder(dynamic rootFolder, string path)
    {
        if (string.IsNullOrEmpty(path)) return rootFolder;
        var parts = path.Replace("\\", "/").Split('/');
        dynamic current = rootFolder;
        foreach (var part in parts)
        {
            if (string.IsNullOrEmpty(part)) continue;
            try
            {
                dynamic subItem = FindItemInFolder(current, part);
                if (subItem == null) return null;
                try
                {
                    if (!subItem.IsFolder) return null;
                    dynamic subFolder = subItem.GetFolder;
                    if (subFolder == null) return null;
                    if (current != rootFolder) try { Marshal.FinalReleaseComObject(current); } catch { }
                    current = subFolder;
                    try { Marshal.FinalReleaseComObject(subItem); } catch { }
                }
                catch { try { Marshal.FinalReleaseComObject(subItem); } catch { } return null; }
            }
            catch { return null; }
        }
        return current;
    }

    static void CopyFile(string selector, string remote, string local)
    {
        var device = FindDevice(selector);
        if (device == null) return;
        try
        {
            var d = device.Value;
            local = Path.GetFullPath(local);
            Directory.CreateDirectory(Path.GetDirectoryName(local) ?? ".");
            dynamic shell = null, computer = null, deviceItem = null, deviceFolder = null, srcFolder = null, srcItem = null;
            try
            {
                shell = GetShell();
                computer = shell.NameSpace(17);
                deviceItem = computer.Items().Item(d.Index);
                if (deviceItem == null) { Console.WriteLine("Cannot access device"); return; }
                deviceFolder = deviceItem.GetFolder;
                if (deviceFolder == null) { Console.WriteLine("Cannot open device"); return; }
                string remoteDir = "";
                string remoteFile = remote.Replace("\\", "/");
                int lastSlash = remoteFile.LastIndexOf('/');
                if (lastSlash >= 0) { remoteDir = remoteFile.Substring(0, lastSlash); remoteFile = remoteFile.Substring(lastSlash + 1); }
                srcFolder = string.IsNullOrEmpty(remoteDir) ? deviceFolder : NavigateToFolder(deviceFolder, remoteDir);
                if (srcFolder == null) { Console.WriteLine($"Remote directory '{remoteDir}' not found"); return; }
                srcItem = FindItemInFolder(srcFolder, remoteFile);
                if (srcItem == null) { Console.WriteLine($"File '{remoteFile}' not found"); return; }
                Console.WriteLine($"Copying: {d.Name}:{remote}");
                Console.WriteLine($"  To: {local}");
                bool copied = CopyViaIStream(srcItem, local);
                if (copied) { var fi = new FileInfo(local); Console.WriteLine($"  Done! {FormatSize(fi.Length)} copied."); }
                else
                {
                    Console.Write("  Trying PowerShell fallback...");
                    copied = CopyViaPowerShell(d.Index, remote, local);
                    Console.WriteLine(copied ? " done!" : " failed.");
                }
                if (!copied) Console.WriteLine("  Copy failed. Try manually via Explorer.");
            }
            finally
            {
                if (srcItem != null) try { Marshal.ReleaseComObject(srcItem); } catch { }
                if (srcFolder != null && srcFolder != deviceFolder) try { Marshal.ReleaseComObject(srcFolder); } catch { }
                if (deviceFolder != null) try { Marshal.ReleaseComObject(deviceFolder); } catch { }
                if (deviceItem != null) try { Marshal.ReleaseComObject(deviceItem); } catch { }
                if (computer != null) try { Marshal.ReleaseComObject(computer); } catch { }
                if (shell != null) try { Marshal.ReleaseComObject(shell); } catch { }
            }
        }
        catch (Exception ex) { Console.WriteLine($"Copy failed: {ex.Message}"); }
    }

    static bool CopyViaPowerShell(int deviceIndex, string remotePath, string localPath)
    {
        remotePath = remotePath.Replace("\\", "/");
        string remoteDir = "";
        string remoteFile = remotePath;
        int lastSlash = remotePath.LastIndexOf('/');
        if (lastSlash >= 0) { remoteDir = remotePath.Substring(0, lastSlash); remoteFile = remotePath.Substring(lastSlash + 1); }
        string script = $@"
$idx = {deviceIndex}
$rdir = '{remoteDir}'
$rfile = '{remoteFile}'
$local = '{localPath}'
try {{
    $shell = New-Object -ComObject Shell.Application
    $comp = $shell.NameSpace(17)
    $dev = $comp.Items().Item($idx)
    function Find-FolderItem($fld, $name) {{
        foreach ($it in $fld.Items()) {{ try {{ if ($it.Name.Trim() -eq $name) {{ return $it }} }} catch {{}} }}
        return $null
    }}
    $cur = $dev.GetFolder()
    if ($rdir) {{
        $parts = $rdir -split '/'
        foreach ($p in $parts) {{ if (-not $p) {{ continue }}
            $sub = Find-FolderItem $cur $p
            if (-not $sub) {{ 'ERR_DIR'; return }}
            $cur = $sub.GetFolder()
        }}
    }}
    $file = Find-FolderItem $cur $rfile
    if (-not $file) {{ 'ERR_FILE'; return }}
    $localDir = Split-Path $local -Parent
    $localFolder = $shell.NameSpace($localDir)
    'COPYING'
    $localFolder.CopyHere($file, 20)
    $timeout = [DateTime]::Now.AddSeconds(90)
    while (-not (Test-Path $local) -and [DateTime]::Now -lt $timeout) {{ Start-Sleep -Milliseconds 300 }}
    if (Test-Path $local) {{ 'OK' }} else {{ 'TIMEOUT' }}
}} catch {{ 'ERR_' + $_.Exception.Message }}
";
        string tempFile = Path.GetTempFileName() + ".ps1";
        try
        {
            File.WriteAllText(tempFile, script, System.Text.Encoding.Unicode);
            using var proc = new Process();
            proc.StartInfo = new ProcessStartInfo
            {
                FileName = "powershell",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{tempFile}\"",
                UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8
            };
            proc.Start();
            string output = proc.StandardOutput.ReadToEnd().Trim();
            string error = proc.StandardError.ReadToEnd().Trim();
            proc.WaitForExit(120000);
            if (output.Contains("OK")) { var fi = new FileInfo(localPath); if (fi.Exists) Console.Write($" Done! {FormatSize(fi.Length)}"); return true; }
            if (output.Contains("TIMEOUT")) return File.Exists(localPath);
            if (output.Contains("ERR_DIR")) { Console.Write(" Remote directory not found"); return false; }
            if (output.Contains("ERR_FILE")) { Console.Write(" Remote file not found"); return false; }
            if (!string.IsNullOrEmpty(output)) Console.Write($" Result: {output}");
            if (!string.IsNullOrEmpty(error)) Console.Write($" Error: {error}");
            return false;
        }
        finally { try { File.Delete(tempFile); } catch { } }
    }

    static string FormatSize(long bytes)
    {
        string[] suf = { "B", "KB", "MB", "GB", "TB" };
        int i = 0; double d = bytes;
        while (d >= 1024 && i < suf.Length - 1) { d /= 1024; i++; }
        return $"{d:0.##} {suf[i]}";
    }

    static string NormalizePath(string path)
    {
        if (string.IsNullOrEmpty(path)) return "/";
        return "/" + path.Replace("\\", "/").Trim('/');
    }

    // Shell COM interfaces for IStream copy
    [ComImport, Guid("0000000c-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IStream
    {
        [PreserveSig] int Read([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, IntPtr pcbRead);
        [PreserveSig] int Write([MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, uint cb, IntPtr pcbWritten);
        [PreserveSig] int Seek(long dlibMove, uint dwOrigin, IntPtr plibNewPosition);
        void SetSize(long libNewSize);
        [PreserveSig] int CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten);
        void Commit(uint grfCommitFlags);
        void Revert();
        void LockRegion(long libOffset, long cb, uint dwLockType);
        void UnlockRegion(long libOffset, long cb, uint dwLockType);
        void Stat(out STATSTG pstatstg, uint grfStatFlag);
        void Clone(out IStream ppstm);
    }

    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItem
    {
        [PreserveSig] int BindToHandler(object pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        [PreserveSig] int GetParent(out IShellItem ppsi);
        [PreserveSig] int GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        [PreserveSig] int GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    }

    static readonly Guid BHID_Stream = new("1cebb3ab-7c10-499a-a417-92ca16c4cb83");
    static readonly Guid IID_IShellItem = new("43826d1e-e718-42ee-bc55-a1e261c37bfe");
    static readonly Guid IID_IStream = new("0000000c-0000-0000-c000-000000000046");

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHGetIDListFromObject(IntPtr punk, out IntPtr ppidl);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHCreateItemFromIDList(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

    static bool CopyViaIStream(dynamic srcItem, string localPath)
    {
        try
        {
            IntPtr pUnk = Marshal.GetIUnknownForObject(srcItem);
            int hr = SHGetIDListFromObject(pUnk, out IntPtr pidl);
            Marshal.Release(pUnk);
            if (hr != 0 || pidl == IntPtr.Zero) return false;
            try
            {
                hr = SHCreateItemFromIDList(pidl, IID_IShellItem, out IShellItem shellItem);
                if (hr != 0 || shellItem == null) return false;
                try
                {
                    hr = shellItem.BindToHandler(null, BHID_Stream, IID_IStream, out IntPtr streamPtr);
                    if (hr != 0 || streamPtr == IntPtr.Zero) return false;
                    IStream stream = (IStream)Marshal.GetTypedObjectForIUnknown(streamPtr, typeof(IStream));
                    try
                    {
                        using var fs = new FileStream(localPath, FileMode.Create, FileAccess.Write);
                        byte[] buffer = new byte[65536];
                        while (true)
                        {
                            IntPtr bytesRead = Marshal.AllocCoTaskMem(4);
                            try
                            {
                                hr = stream.Read(buffer, (uint)buffer.Length, bytesRead);
                                uint read = (uint)Marshal.ReadInt32(bytesRead);
                                if (read == 0) break;
                                fs.Write(buffer, 0, (int)read);
                            }
                            finally { Marshal.FreeCoTaskMem(bytesRead); }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(stream);
                        Marshal.Release(streamPtr);
                    }
                }
                finally { Marshal.ReleaseComObject(shellItem); }
            }
            finally { Marshal.FreeCoTaskMem(pidl); }
            return true;
        }
        catch { return false; }
    }
}
