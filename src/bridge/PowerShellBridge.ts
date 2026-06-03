import { execFile } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import type { BridgeClient, VbComponentInfo, VbComponentType, WorkbookRef } from './types';

const COMPONENT_TYPE_MAP: Record<number, VbComponentType> = {
  1: 'StdModule',
  2: 'ClassModule',
  3: 'MSForm',
  11: 'ActiveXDesigner',
  100: 'Document'
};

// Run a PowerShell 5.1 script and return trimmed stdout as a string.
// PowerShell is always present on Windows 7+; no native compilation required.
function runPS(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-OutputFormat', 'Text', '-Command', script],
      { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err) reject(new Error(stderr?.trim() || err.message));
        else resolve(stdout.trim());
      }
    );
  });
}

// Escape single-quotes for embedding a value inside a PS single-quoted string.
const escPS = (s: string) => s.replace(/'/g, "''");

// Inline C# that enumerates the Windows Running Object Table to find ALL
// running Excel.Application instances, not just the most-recently-focused one.
// GetActiveObject('Excel.Application') only returns the last ROT registration,
// which misses any Excel process that isn't currently in the foreground.
const ROT_HELPER_CS = `
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public class RotHelper {
    [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int r, out IRunningObjectTable p);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int r, out IBindCtx p);
    public static List<object> GetExcelApps() {
        var list = new List<object>();
        IRunningObjectTable rot; if (GetRunningObjectTable(0, out rot) != 0) return list;
        IEnumMoniker em; rot.EnumRunning(out em); em.Reset();
        var m = new IMoniker[1]; var f = IntPtr.Zero;
        while (em.Next(1, m, f) == 0) {
            IBindCtx ctx; CreateBindCtx(0, out ctx);
            string name; m[0].GetDisplayName(ctx, null, out name);
            Marshal.ReleaseComObject(ctx);
            if (name == null || !name.StartsWith("!")) continue;
            try {
                object obj; rot.GetObject(m[0], out obj);
                var n = obj.GetType().InvokeMember("Name",
                    System.Reflection.BindingFlags.GetProperty, null, obj, null);
                if (n != null && n.ToString().IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0)
                    list.Add(obj);
            } catch {}
        }
        return list;
    }
}
"@
`;

// Resolve a workbook across ALL running Excel instances via ROT.
const WB_RESOLVER_PS = (fullName: string) => `
${ROT_HELPER_CS}
$_wb = $null
foreach ($__xl in [RotHelper]::GetExcelApps()) {
  for ($__i = 1; $__i -le $__xl.Workbooks.Count; $__i++) {
    if ($__xl.Workbooks.Item($__i).FullName -eq '${escPS(fullName)}') {
      $_wb = $__xl.Workbooks.Item($__i); break
    }
  }
  if ($_wb) { break }
}
if (-not $_wb) { throw 'Workbook not open in Excel: ${escPS(fullName)}' }
`;

export class PowerShellBridge implements BridgeClient {
  async isAvailable(): Promise<boolean> {
    if (process.platform !== 'win32') return false;
    try {
      await runPS('Write-Output ok');
      return true;
    } catch {
      return false;
    }
  }

  async listOpenWorkbooks(): Promise<WorkbookRef[]> {
    const script = `
      [Console]::OutputEncoding = [Text.Encoding]::UTF8
      ${ROT_HELPER_CS}
      $results = [System.Collections.Generic.List[PSCustomObject]]::new()
      foreach ($xl in [RotHelper]::GetExcelApps()) {
        for ($i = 1; $i -le $xl.Workbooks.Count; $i++) {
          $wb = $xl.Workbooks.Item($i)
          $results.Add([PSCustomObject]@{ excelPid = [int]$xl.Hwnd; name = $wb.Name; fullName = $wb.FullName })
        }
      }
      if ($results.Count -gt 0) { ConvertTo-Json -Compress -InputObject @($results) } else { '[]' }
    `;
    const output = await runPS(script);
    const parsed = JSON.parse(output || '[]');
    return (Array.isArray(parsed) ? parsed : [parsed]) as WorkbookRef[];
  }

  async listComponents(workbook: WorkbookRef): Promise<VbComponentInfo[]> {
    const script = `
      [Console]::OutputEncoding = [Text.Encoding]::UTF8
      ${WB_RESOLVER_PS(workbook.fullName)}
      $list = for ($i = 1; $i -le $_wb.VBProject.VBComponents.Count; $i++) {
        $c = $_wb.VBProject.VBComponents.Item($i)
        [PSCustomObject]@{ name = $c.Name; type = [int]$c.Type; lineCount = [int]$c.CodeModule.CountOfLines }
      }
      if ($list) { ConvertTo-Json -Compress -InputObject @($list) } else { '[]' }
    `;
    const output = await runPS(script);
    const parsedComponents = JSON.parse(output || '[]');
    const raw = (Array.isArray(parsedComponents) ? parsedComponents : [parsedComponents]) as { name: string; type: number; lineCount: number }[];
    return raw.map(r => ({
      name: r.name,
      type: COMPONENT_TYPE_MAP[r.type] ?? 'Unknown',
      lineCount: r.lineCount
    }));
  }

  async getComponentSource(workbook: WorkbookRef, componentName: string): Promise<string> {
    const script = `
      [Console]::OutputEncoding = [Text.Encoding]::UTF8
      ${WB_RESOLVER_PS(workbook.fullName)}
      $c = $_wb.VBProject.VBComponents.Item('${escPS(componentName)}')
      $n = [int]$c.CodeModule.CountOfLines
      if ($n -gt 0) { $c.CodeModule.Lines(1, $n) } else { '' }
    `;
    return await runPS(script);
  }

  async setComponentSource(workbook: WorkbookRef, componentName: string, source: string): Promise<void> {
    // Write the new source to a temp file so we never have to escape arbitrary
    // VBA code inside a PowerShell string (dollar signs, backticks, quotes, etc.).
    const tmpFile = path.join(os.tmpdir(), `bassync_${Date.now()}.tmp`);
    fs.writeFileSync(tmpFile, source, 'utf8');
    try {
      const script = `
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        ${WB_RESOLVER_PS(workbook.fullName)}
        $newSource = [IO.File]::ReadAllText('${escPS(tmpFile)}', [Text.Encoding]::UTF8)
        $c = $_wb.VBProject.VBComponents.Item('${escPS(componentName)}')
        $m = $c.CodeModule
        if ($m.CountOfLines -gt 0) { $m.DeleteLines(1, $m.CountOfLines) }
        $m.AddFromString($newSource)
        Write-Output 'ok'
      `;
      await runPS(script);
    } finally {
      try { fs.unlinkSync(tmpFile); } catch { /* ignore */ }
    }
  }

  dispose(): void {
    // No persistent resources; each runPS call is a fresh process.
  }
}
