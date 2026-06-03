# Diagnostic: enumerate ALL running Excel instances via the Running Object Table
[Console]::OutputEncoding = [Text.Encoding]::UTF8

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class RotHelper {
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);
    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    public static List<object> GetExcelApplications() {
        var results = new List<object>();
        IRunningObjectTable rot;
        if (GetRunningObjectTable(0, out rot) != 0) return results;
        IEnumMoniker enumMoniker;
        rot.EnumRunning(out enumMoniker);
        enumMoniker.Reset();
        var moniker = new IMoniker[1];
        var fetched = IntPtr.Zero;
        while (enumMoniker.Next(1, moniker, fetched) == 0) {
            IBindCtx ctx;
            CreateBindCtx(0, out ctx);
            string name;
            moniker[0].GetDisplayName(ctx, null, out name);
            Marshal.ReleaseComObject(ctx);
            // Excel registers its Application as "!{GUID}" monikers;
            // each open workbook registers as a file moniker ending with .xlsm/.xlsx etc.
            // We grab anything that maps to an Excel.Application dispatch object.
            if (name != null && name.StartsWith("!")) {
                try {
                    object obj;
                    rot.GetObject(moniker[0], out obj);
                    // Test if it quacks like Excel.Application
                    var t = obj.GetType();
                    var nameProp = t.InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, obj, null);
                    if (nameProp != null && nameProp.ToString().Contains("Excel"))
                        results.Add(obj);
                } catch { }
            }
        }
        return results;
    }
}
"@

$instances = [RotHelper]::GetExcelApplications()
Write-Output ("Excel instances found via ROT: " + $instances.Count)
foreach ($xl in $instances) {
    Write-Output ("  Application: " + $xl.Name + " | Hwnd: " + $xl.Hwnd + " | Workbooks: " + $xl.Workbooks.Count)
    for ($i = 1; $i -le $xl.Workbooks.Count; $i++) {
        $wb = $xl.Workbooks.Item($i)
        Write-Output ("    - " + $wb.Name + " => " + $wb.FullName)
    }
}
