import type { BridgeClient, VbComponentInfo, VbComponentType, WorkbookRef } from './types';

// winax is a native module; loaded lazily so the extension host can start
// even on non-Windows or before electron-rebuild has been run.
type WinaxModule = {
  Object: (progId: string, options?: { activate?: boolean }) => any;
  getObject: (progId: string) => any;
  release: (obj: any) => void;
};

let winaxModule: WinaxModule | undefined;
let loadError: Error | undefined;

function loadWinax(): WinaxModule {
  if (winaxModule) return winaxModule;
  if (loadError) throw loadError;
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    winaxModule = require('winax') as WinaxModule;
    return winaxModule;
  } catch (err) {
    loadError = err as Error;
    throw loadError;
  }
}

const COMPONENT_TYPE_MAP: Record<number, VbComponentType> = {
  1: 'StdModule',
  2: 'ClassModule',
  3: 'MSForm',
  11: 'ActiveXDesigner',
  100: 'Document'
};

function mapComponentType(raw: number): VbComponentType {
  return COMPONENT_TYPE_MAP[raw] ?? 'Unknown';
}

export class WinaxBridge implements BridgeClient {
  private excelInstances = new Map<number, any>();

  async isAvailable(): Promise<boolean> {
    if (process.platform !== 'win32') return false;
    try {
      loadWinax();
      return true;
    } catch {
      return false;
    }
  }

  async listOpenWorkbooks(): Promise<WorkbookRef[]> {
    const winax = loadWinax();
    // GetActiveObject only returns the most-recently-registered Excel in the
    // Running Object Table. That is enough for the MVP; a multi-instance
    // enumerator over the ROT is a follow-up.
    let excel: any;
    try {
      excel = winax.getObject('Excel.Application');
    } catch {
      return [];
    }

    const pid = this.resolvePid(excel);
    this.excelInstances.set(pid, excel);

    const results: WorkbookRef[] = [];
    const workbooks = excel.Workbooks;
    const count: number = workbooks.Count;
    for (let i = 1; i <= count; i++) {
      const wb = workbooks.Item(i);
      results.push({
        excelPid: pid,
        fullName: String(wb.FullName),
        name: String(wb.Name)
      });
    }
    return results;
  }

  async listComponents(workbook: WorkbookRef): Promise<VbComponentInfo[]> {
    const wb = this.resolveWorkbook(workbook);
    const project = wb.VBProject;
    const components = project.VBComponents;
    const count: number = components.Count;
    const results: VbComponentInfo[] = [];
    for (let i = 1; i <= count; i++) {
      const c = components.Item(i);
      const code = c.CodeModule;
      results.push({
        name: String(c.Name),
        type: mapComponentType(Number(c.Type)),
        lineCount: Number(code.CountOfLines)
      });
    }
    return results;
  }

  async getComponentSource(workbook: WorkbookRef, componentName: string): Promise<string> {
    const wb = this.resolveWorkbook(workbook);
    const component = wb.VBProject.VBComponents.Item(componentName);
    const code = component.CodeModule;
    const lineCount: number = Number(code.CountOfLines);
    if (lineCount <= 0) return '';
    return String(code.Lines(1, lineCount));
  }

  async setComponentSource(workbook: WorkbookRef, componentName: string, source: string): Promise<void> {
    const wb = this.resolveWorkbook(workbook);
    const component = wb.VBProject.VBComponents.Item(componentName);
    const code = component.CodeModule;
    const lineCount: number = Number(code.CountOfLines);
    if (lineCount > 0) code.DeleteLines(1, lineCount);
    code.AddFromString(source);
  }

  dispose(): void {
    this.excelInstances.clear();
  }

  private resolvePid(excel: any): number {
    try {
      return Number(excel.Hwnd);
    } catch {
      return 0;
    }
  }

  private resolveWorkbook(ref: WorkbookRef): any {
    const winax = loadWinax();
    const excel = this.excelInstances.get(ref.excelPid) ?? winax.getObject('Excel.Application');
    const workbooks = excel.Workbooks;
    const count: number = workbooks.Count;
    for (let i = 1; i <= count; i++) {
      const wb = workbooks.Item(i);
      if (String(wb.FullName) === ref.fullName) return wb;
    }
    throw new Error(`Workbook not currently open in Excel: ${ref.fullName}`);
  }
}
