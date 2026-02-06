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


      console.log(html);
      return html;
}