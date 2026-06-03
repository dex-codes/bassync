import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { PowerShellBridge } from './bridge/PowerShellBridge';
import type { BassyncManifest, BridgeClient, VbComponentType, WorkbookRef } from './bridge/types';

// File extension for each VBA component type.
// MSForm and ActiveXDesigner are skipped — .frx is binary and cannot be round-tripped as text.
const EXT_MAP: Partial<Record<VbComponentType, string>> = {
  StdModule: '.bas',
  ClassModule: '.cls',
  Document: '.cls',
};

let bridge: BridgeClient | undefined;

function getBridge(): BridgeClient {
  if (!bridge) bridge = new PowerShellBridge();
  return bridge;
}

async function ensureAvailable(): Promise<boolean> {
  if (await getBridge().isAvailable()) return true;
  vscode.window.showErrorMessage(
    'Bassync requires Windows with PowerShell 5.1+ and Excel installed.'
  );
  return false;
}

async function pickWorkbook(): Promise<WorkbookRef | undefined> {
  const workbooks = await getBridge().listOpenWorkbooks();
  if (workbooks.length === 0) {
    vscode.window.showWarningMessage(
      'No open Excel workbooks were found. Open the .xlsm in Excel first.'
    );
    return undefined;
  }
  const pick = await vscode.window.showQuickPick(
    workbooks.map((w) => ({ label: w.name, description: w.fullName, workbook: w })),
    { placeHolder: 'Select a workbook to attach to' }
  );
  return pick?.workbook;
}

/** Derive the mirror folder path from the workbook's full path.
 *  e.g. C:\Docs\MyBook.xlsm  →  C:\Docs\MyBook.bassync
 */
function mirrorDir(workbook: WorkbookRef): string {
  const dir = path.dirname(workbook.fullName);
  const base = path.basename(workbook.fullName, path.extname(workbook.fullName));
  return path.join(dir, `${base}.bassync`);
}

async function attachToWorkbookCommand(): Promise<void> {
  if (!(await ensureAvailable())) return;

  const workbook = await pickWorkbook();
  if (!workbook) return;

  await vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: `Bassync: pulling "${workbook.name}"…`,
      cancellable: false,
    },
    async (progress) => {
      const b = getBridge();

      progress.report({ message: 'Listing components…' });
      const components = await b.listComponents(workbook);

      const textComponents = components.filter((c) => EXT_MAP[c.type] !== undefined);
      const skipped = components.length - textComponents.length;

      const outDir = mirrorDir(workbook);
      fs.mkdirSync(outDir, { recursive: true });

      // Write manifest so the mirror window knows which workbook to push back to.
      const manifest: BassyncManifest = {
        workbookFullName: workbook.fullName,
        workbookName: workbook.name,
        pulledAt: new Date().toISOString(),
      };
      fs.writeFileSync(path.join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2), 'utf8');

      let written = 0;
      for (const c of textComponents) {
        progress.report({
          message: `Writing ${c.name}… (${written + 1}/${textComponents.length})`,
          increment: 100 / textComponents.length,
        });
        const source = await b.getComponentSource(workbook, c.name);
        const filePath = path.join(outDir, `${c.name}${EXT_MAP[c.type]}`);
        fs.writeFileSync(filePath, source, 'utf8');
        written++;
      }

      // Brief summary shown in the notification before it auto-dismisses.
      const skipMsg = skipped > 0 ? ` (${skipped} form/designer skipped)` : '';
      progress.report({ message: `Done — ${written} file(s) written${skipMsg}` });

      // Open the mirror folder as a NEW VSCode window.
      const folderUri = vscode.Uri.file(outDir);
      await vscode.commands.executeCommand('vscode.openFolder', folderUri, {
        forceNewWindow: true,
      });
    }
  );
}

async function listOpenWorkbooksCommand(): Promise<void> {
  if (!(await ensureAvailable())) return;
  try {
    const workbooks = await getBridge().listOpenWorkbooks();
    const out = vscode.window.createOutputChannel('Bassync');
    out.clear();
    out.appendLine(`Found ${workbooks.length} open workbook(s):`);
    for (const wb of workbooks) {
      out.appendLine(`  - ${wb.name}  ${wb.fullName}`);
    }
    out.show(true);
  } catch (err) {
    vscode.window.showErrorMessage(`Bassync: ${(err as Error).message}`);
  }
}

function setupMirrorWatcher(context: vscode.ExtensionContext, mirrorFolder: string): void {
  const manifestPath = path.join(mirrorFolder, 'manifest.json');
  let manifest: BassyncManifest;
  try {
    manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8')) as BassyncManifest;
  } catch {
    return; // not a valid bassync workspace
  }

  const workbook: WorkbookRef = {
    fullName: manifest.workbookFullName,
    name: manifest.workbookName,
    excelPid: 0,
  };

  // Status bar item so the user always knows which workbook is live.
  const statusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 10);
  statusBar.text = `$(sync) Bassync: ${manifest.workbookName}`;
  statusBar.tooltip = `Live-syncing to ${manifest.workbookFullName}`;
  statusBar.show();
  context.subscriptions.push(statusBar);

  context.subscriptions.push(
    vscode.workspace.onDidSaveTextDocument(async (doc) => {
      const filePath = doc.uri.fsPath;
      // Only handle files inside this mirror folder.
      if (!filePath.startsWith(mirrorFolder + path.sep)) return;

      const ext = path.extname(filePath);
      if (ext !== '.bas' && ext !== '.cls') return;

      const componentName = path.basename(filePath, ext);
      statusBar.text = `$(sync~spin) Bassync: pushing ${componentName}…`;

      try {
        await getBridge().setComponentSource(workbook, componentName, doc.getText());
        statusBar.text = `$(check) Bassync: ${manifest.workbookName}`;
        vscode.window.setStatusBarMessage(`Bassync: pushed ${componentName} to Excel ✓`, 3000);
      } catch (err) {
        statusBar.text = `$(error) Bassync: ${manifest.workbookName}`;
        vscode.window.showErrorMessage(
          `Bassync: failed to push "${componentName}" — ${(err as Error).message}`
        );
      } finally {
        // Reset icon after a moment if no error replaced it.
        setTimeout(() => {
          if (statusBar.text.startsWith('$(check)')) {
            statusBar.text = `$(sync) Bassync: ${manifest.workbookName}`;
          }
        }, 3000);
      }
    })
  );
}

export function activate(context: vscode.ExtensionContext): void {
  context.subscriptions.push(
    vscode.commands.registerCommand('bassync.attachToWorkbook', attachToWorkbookCommand),
    vscode.commands.registerCommand('bassync.listOpenWorkbooks', listOpenWorkbooksCommand)
  );

  // If this VSCode window was opened on a .bassync mirror folder, start watching for saves.
  for (const folder of vscode.workspace.workspaceFolders ?? []) {
    const manifestPath = path.join(folder.uri.fsPath, 'manifest.json');
    if (fs.existsSync(manifestPath)) {
      setupMirrorWatcher(context, folder.uri.fsPath);
    }
  }
}

export function deactivate(): void {
  bridge?.dispose();
  bridge = undefined;
}
