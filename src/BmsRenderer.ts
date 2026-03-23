import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';

export function openBmsRenderer(context: vscode.ExtensionContext, document: vscode.TextDocument) {
    const panel = vscode.window.createWebviewPanel(
        'bmsRenderer',
        'BMS Renderer',
        vscode.ViewColumn.Beside,
        {
            enableScripts: true
        }
    );

    const bmsSource = document.getText();
    const fileName = document.uri.path.split('/').pop()?.replace(/\.bms$/i, '') ?? 'BMS';

    panel.webview.html = getWebviewContent(context, fileName, bmsSource);

    // Handle messages from the webview
    panel.webview.onDidReceiveMessage(async message => {
        if (message.command === 'revealField') {
            revealFieldInDocument(document, message.fieldId);
        } else if (message.command === 'saveBms') {
            await vscode.workspace.fs.writeFile(
                document.uri,
                Buffer.from(message.content as string, 'utf8')
            );
            vscode.window.showInformationMessage(
                `Saved: ${path.basename(document.uri.fsPath)}`
            );
        } else if (message.command === 'syncBms') {
            const edit = new vscode.WorkspaceEdit();
            const fullRange = new vscode.Range(
                document.positionAt(0),
                document.positionAt(document.getText().length)
            );
            edit.replace(document.uri, fullRange, message.content as string);
            await vscode.workspace.applyEdit(edit);
        }
    }, undefined, context.subscriptions);
}

function revealFieldInDocument(document: vscode.TextDocument, fieldId: string) {
    const text  = document.getText();
    const lines = text.split('\n');
    // Find the line where fieldId appears as a DFHMDF label (first column token)
    const labelRe = new RegExp(`^(${escapeRegex(fieldId)})\\s+DFHMDF\\b`, 'im');
    for (let i = 0; i < lines.length; i++) {
        if (labelRe.test(lines[i])) {
            const range = new vscode.Range(i, 0, i, lines[i].length);
            vscode.window.showTextDocument(document, {
                viewColumn: vscode.ViewColumn.One,
                selection: range,
                preserveFocus: false,
            });
            return;
        }
    }
    vscode.window.showInformationMessage(`Field "${fieldId}" not found in source.`);
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getWebviewContent(context: vscode.ExtensionContext, fileName: string, bmsSource: string): string {
    const htmlPath = path.join(
        context.extensionPath,
        'src',
        'media',
        'renderer.html'
      );

      let html = fs.readFileSync(htmlPath, 'utf8');
      html = html
        .replaceAll('{{TITLE}}', fileName)
        .replace('"{{SOURCE}}"', JSON.stringify(bmsSource));

      return html;
}