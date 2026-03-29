import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';

export function openBmsRenderer(context: vscode.ExtensionContext, document: vscode.TextDocument) {
    const panel = vscode.window.createWebviewPanel(
        'bmsRenderer',
        'BMS Renderer',
        vscode.ViewColumn.Beside,
        {
            enableScripts: true,
            localResourceRoots: [
                vscode.Uri.joinPath(context.extensionUri, 'src', 'media')
            ]
        }
    );

    const bmsSource = document.getText();
    const fileName = document.uri.path.split('/').pop()?.replace(/\.bms$/i, '') ?? 'BMS';

    // Load persisted UI config
    const savedConfig = {
        fill:       context.workspaceState.get<string>('bmsRenderer.fill',       'empty'),
        sync:       context.workspaceState.get<boolean>('bmsRenderer.sync',      false),
        autoResize: context.workspaceState.get<boolean>('bmsRenderer.autoResize',false),
        theme:      context.globalState.get<string>('bmsRenderer.theme',         'dark'),
    };

    panel.webview.html = getWebviewContent(context, panel.webview, fileName, bmsSource, savedConfig);

    // Prevent echo-back when we ourselves apply a WorkspaceEdit
    let applyingSync = false;

    // Push file changes into the renderer immediately
    const changeDisposable = vscode.workspace.onDidChangeTextDocument(e => {
        if (e.document.uri.toString() !== document.uri.toString()) return;
        if (applyingSync) return;
        panel.webview.postMessage({ command: 'updateSource', source: e.document.getText() });
    });
    panel.onDidDispose(() => changeDisposable.dispose());

    // Handle messages from the webview
    panel.webview.onDidReceiveMessage(async message => {
        if (message.command === 'revealField') {
            revealFieldInDocument(document, message.fieldId);
        } else if (message.command === 'saveConfig') {
            // Persist per-workspace settings
            if (message.fill       !== undefined) { await context.workspaceState.update('bmsRenderer.fill',       message.fill); }
            if (message.sync       !== undefined) { await context.workspaceState.update('bmsRenderer.sync',       message.sync); }
            if (message.autoResize !== undefined) { await context.workspaceState.update('bmsRenderer.autoResize', message.autoResize); }
            // Persist global setting
            if (message.theme      !== undefined) { await context.globalState.update('bmsRenderer.theme',         message.theme); }
        } else if (message.command === 'saveBms') {
            await vscode.workspace.fs.writeFile(
                document.uri,
                Buffer.from(message.content as string, 'utf8')
            );
            vscode.window.showInformationMessage(
                `Saved: ${path.basename(document.uri.fsPath)}`
            );
        } else if (message.command === 'syncBms') {
            const newContent = message.content as string;
            if (document.getText() === newContent) return;
            applyingSync = true;
            try {
                const edit = new vscode.WorkspaceEdit();
                const fullRange = new vscode.Range(
                    document.positionAt(0),
                    document.positionAt(document.getText().length)
                );
                edit.replace(document.uri, fullRange, newContent);
                await vscode.workspace.applyEdit(edit);
            } finally {
                applyingSync = false;
            }
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

function getWebviewContent(
    context: vscode.ExtensionContext,
    webview: vscode.Webview,
    fileName: string,
    bmsSource: string,
    savedConfig: { fill: string; sync: boolean; autoResize: boolean; theme: string }
): string {
    const htmlPath = path.join(context.extensionPath, 'src', 'media', 'renderer.html');

    const cssUri = webview.asWebviewUri(
        vscode.Uri.joinPath(context.extensionUri, 'src', 'media', 'renderer.css')
    );
    const jsUri = webview.asWebviewUri(
        vscode.Uri.joinPath(context.extensionUri, 'src', 'media', 'renderer.js')
    );

    let html = fs.readFileSync(htmlPath, 'utf8');
    html = html
        .replaceAll('{{TITLE}}', fileName)
        .replace('"{{SOURCE}}"', JSON.stringify(bmsSource))
        .replace('>{{CONFIG}}<', `>${JSON.stringify(savedConfig)}<`)
        .replace('{{CSS_URI}}', cssUri.toString())
        .replace('{{JS_URI}}', jsUri.toString());

    return html;
}