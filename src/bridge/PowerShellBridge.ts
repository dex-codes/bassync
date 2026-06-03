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

  async addComponent(workbook: WorkbookRef, componentName: string, vbType: number, source: string): Promise<void> {
    const tmpFile = path.join(os.tmpdir(), `bassync_${Date.now()}.tmp`);
    fs.writeFileSync(tmpFile, source, 'utf8');
    try {
      const script = `
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        ${WB_RESOLVER_PS(workbook.fullName)}
        $newSource = [IO.File]::ReadAllText('${escPS(tmpFile)}', [Text.Encoding]::UTF8)
        $c = $_wb.VBProject.VBComponents.Add(${vbType})
        $c.Name = '${escPS(componentName)}'
        $m = $c.CodeModule
        if ($m.CountOfLines -gt 0) { $m.DeleteLines(1, $m.CountOfLines) }
        if ($newSource.Trim() -ne '') { $m.AddFromString($newSource) }
        $_wb.Save()
        Write-Output 'ok'
      `;
      await runPS(script);
    } finally {
      try { fs.unlinkSync(tmpFile); } catch { /* ignore */ }
    }
  }

  async removeComponent(workbook: WorkbookRef, componentName: string): Promise<void> {
    const script = `
      [Console]::OutputEncoding = [Text.Encoding]::UTF8
      ${WB_RESOLVER_PS(workbook.fullName)}
      $c = $null
      try { $c = $_wb.VBProject.VBComponents.Item('${escPS(componentName)}') } catch {}
      if ($c) {
        $_wb.VBProject.VBComponents.Remove($c)
        $_wb.Save()
      }
      Write-Output 'ok'
    `;
    await runPS(script);
  }

  dispose(): void {
    // No persistent resources; each runPS call is a fresh process.
  }
}

// ── Ribbon XML helpers ────────────────────────────────────────────────────────
// These work directly on the .xlsm ZIP and do not go through COM.
// They are standalone exports so extension.ts can call them without a
// BridgeClient instance (ribbon is file-based, not VBProject-based).

export const RIBBON_TEMPLATE = `<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="customTab" label="Custom">
        <group id="customGroup" label="Actions">
          <!-- Add ribbon controls here -->
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>`;

/**
 * Read the ribbon customUI XML from inside the .xlsm ZIP.
 * Returns an empty string if no customUI is defined.
 * Safe to call while Excel has the workbook open.
 */
export async function getRibbonXml(xlsmPath: string): Promise<string> {
  const script = `
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    Add-Type -Assembly 'System.IO.Compression' -ErrorAction SilentlyContinue
    Add-Type -Assembly 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
    try {
      $z = [IO.Compression.ZipFile]::OpenRead('${escPS(xlsmPath)}')
      $e = $z.Entries | Where-Object {
        $_.FullName -eq 'customUI14/customUI14.xml' -or $_.FullName -eq 'customUI/customUI.xml'
      } | Select-Object -First 1
      if ($e) {
        $r = [IO.StreamReader]::new($e.Open(), [Text.Encoding]::UTF8)
        $c = $r.ReadToEnd(); $r.Close(); $z.Dispose(); Write-Output $c
      } else { $z.Dispose(); Write-Output '' }
    } catch { Write-Output '' }
  `;
  return await runPS(script);
}

/**
 * Write ribbon customUI XML back into the .xlsm ZIP.
 *
 * Because Excel locks the file while the workbook is open, this function:
 *   1. Finds the workbook via the Running Object Table
 *   2. Saves and closes it in Excel
 *   3. Modifies the ZIP on disk (update or add customUI14/customUI14.xml
 *      plus the _rels/.rels relationship entry if it did not exist before)
 *   4. Reopens the workbook in Excel
 *
 * If the workbook is not currently open in Excel, steps 1/2/4 are skipped
 * and the ZIP is modified directly.
 */
export async function setRibbonXml(xlsmPath: string, xml: string): Promise<void> {
  const tmpFile = path.join(os.tmpdir(), `bassync_ribbon_${Date.now()}.tmp`);
  fs.writeFileSync(tmpFile, xml, 'utf8');
  try {
    const script = `
      [Console]::OutputEncoding = [Text.Encoding]::UTF8
      Add-Type -Assembly 'System.IO.Compression' -ErrorAction SilentlyContinue
      Add-Type -Assembly 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
      ${ROT_HELPER_CS}

      $r_path = '${escPS(xlsmPath)}'
      $r_xml  = [IO.File]::ReadAllText('${escPS(tmpFile)}', [Text.Encoding]::UTF8)
      $r_xl   = $null
      $r_open = $false

      # --- 1. Save + close in Excel if the workbook is currently open ----------
      foreach ($r_app in [RotHelper]::GetExcelApps()) {
        for ($r_i = 1; $r_i -le $r_app.Workbooks.Count; $r_i++) {
          $r_wb = $r_app.Workbooks.Item($r_i)
          if ($r_wb.FullName -eq $r_path) {
            $r_xl   = $r_app
            $r_wb.Save()
            $r_wb.Close($false)
            $r_open = $true
            break
          }
        }
        if ($r_open) { break }
      }
      if ($r_open) { Start-Sleep -Milliseconds 400 }

      # --- 2. Modify the ZIP ---------------------------------------------------
      $r_zip     = [IO.Compression.ZipFile]::Open($r_path, [IO.Compression.ZipArchiveMode]::Update)
      $r_cuiPath = 'customUI14/customUI14.xml'
      $r_prev    = $r_zip.Entries | Where-Object {
        $_.FullName -eq 'customUI14/customUI14.xml' -or $_.FullName -eq 'customUI/customUI.xml'
      } | Select-Object -First 1
      $r_isNew = $false

      if ($r_prev) {
        $r_cuiPath = $r_prev.FullName
        $r_prev.Delete()           # delete then recreate — only safe way to replace in-place
      } else {
        $r_isNew = $true
      }

      $r_entry  = $r_zip.CreateEntry($r_cuiPath)
      $r_stream = $r_entry.Open()
      $r_bytes  = [Text.Encoding]::UTF8.GetBytes($r_xml)
      $r_stream.Write($r_bytes, 0, $r_bytes.Length)
      $r_stream.Close()

      # If this is brand-new ribbon XML, wire up the package relationship so Excel recognises it
      if ($r_isNew) {
        $r_relsEntry = $r_zip.Entries | Where-Object { $_.FullName -eq '_rels/.rels' } | Select-Object -First 1
        if ($r_relsEntry) {
          $r_reader  = [IO.StreamReader]::new($r_relsEntry.Open(), [Text.Encoding]::UTF8)
          $r_relsXml = $r_reader.ReadToEnd(); $r_reader.Close()
          if ($r_relsXml -notmatch 'ui/extensibility') {
            $r_relId   = 'rId' + [guid]::NewGuid().ToString('N').Substring(0, 8)
            $r_relsXml = $r_relsXml -replace '</Relationships>',
              "<Relationship Id=""$r_relId"" Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Target=""$r_cuiPath""/></Relationships>"
            $r_relsEntry.Delete()
            $r_newRels = $r_zip.CreateEntry('_rels/.rels')
            $r_rw = [IO.StreamWriter]::new($r_newRels.Open(), [Text.Encoding]::UTF8)
            $r_rw.Write($r_relsXml); $r_rw.Close()
          }
        }
      }
      $r_zip.Dispose()

      # --- 3. Reopen in Excel if it was open -----------------------------------
      if ($r_open -and $r_xl) { $r_xl.Workbooks.Open($r_path) | Out-Null }
      Write-Output 'ok'
    `;
    await runPS(script);
  } finally {
    try { fs.unlinkSync(tmpFile); } catch { /* ignore */ }
  }
}
