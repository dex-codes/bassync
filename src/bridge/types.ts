export type VbComponentType =
  | 'StdModule'
  | 'ClassModule'
  | 'MSForm'
  | 'Document'
  | 'ActiveXDesigner'
  | 'Unknown';

export interface WorkbookRef {
  excelPid: number;
  fullName: string;
  name: string;
}

export interface VbComponentInfo {
  name: string;
  type: VbComponentType;
  lineCount: number;
}

export interface BridgeClient {
  isAvailable(): Promise<boolean>;
  listOpenWorkbooks(): Promise<WorkbookRef[]>;
  listComponents(workbook: WorkbookRef): Promise<VbComponentInfo[]>;
  getComponentSource(workbook: WorkbookRef, componentName: string): Promise<string>;
  setComponentSource(workbook: WorkbookRef, componentName: string, source: string): Promise<void>;
  dispose(): void;
}

export interface BassyncManifest {
  workbookFullName: string;
  workbookName: string;
  pulledAt: string;
}
