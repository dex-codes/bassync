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
  /** Add a brand-new VBA component (vbType 1=StdModule, 2=ClassModule) and save the workbook. */
  addComponent(workbook: WorkbookRef, componentName: string, vbType: number, source: string): Promise<void>;
  /** Remove a VBA component by name and save the workbook. */
  removeComponent(workbook: WorkbookRef, componentName: string): Promise<void>;
  dispose(): void;
}

export interface BassyncManifest {
  workbookFullName: string;
  workbookName: string;
  pulledAt: string;
}
