import * as vscode from 'vscode';

const BMS_MACROS = [
    'DFHMSD',
    'DFHMDI',
    'DFHMDF'
];
  
const BMS_ATTRIBUTES = [
    'TYPE', 'MODE', 'LANG', 'STORAGE', 'CTRL', 'EXTATT', 'TERM',
    'TIOAPFX', 'MAPATTS', 'DSATTS', 'SIZE', 'COLUMN', 'LINE',
    'POS', 'LENGTH', 'INITIAL', 'ATTRB', 'HILIGHT', 'COLOR', 'OUTLINE'
];
  
const BMS_CONSTANTS = [
    'YES', 'NO', 'ON', 'OFF', 'AUTO', 'COBOL', 'INOUT',
    'ASKIP', 'NORM', 'PROT', 'UNPROT', 'BRT',
    'FREEKB', 'FINAL', '3270-2'
];

const BMS_COLORS = [
    'BLUE', 'GREEN', 'WHITE', 'CYAN', 'RED', 'YELLOW', 'MAGENTA', 'BLACK'
];

export class BmsCompletionProvider implements vscode.CompletionItemProvider {

    provideCompletionItems(
      document: vscode.TextDocument,
      position: vscode.Position
    ): vscode.ProviderResult<vscode.CompletionItem[]> {
  
      const line = document.lineAt(position.line).text;
      const col = position.character;
  
      const items: vscode.CompletionItem[] = [];
  
      // ───────────────────────────────
      // MACROS → only after column 8
      // ───────────────────────────────
      if (col >= 8) {
        for (const macro of BMS_MACROS) {
          const item = new vscode.CompletionItem(
            macro,
            vscode.CompletionItemKind.Keyword
          );
          item.insertText = macro;
          item.detail = 'BMS Macro';
          items.push(item);
        }
      }
  
      // ───────────────────────────────
      // ATTRIBUTES → anywhere after macro
      // ───────────────────────────────
      for (const attr of BMS_ATTRIBUTES) {
        const item = new vscode.CompletionItem(
          attr,
          vscode.CompletionItemKind.Field
        );
        item.insertText = attr + '=';
        item.detail = 'BMS Attribute';
        items.push(item);
      }
  
      // ───────────────────────────────
      // CONSTANTS → values
      // ───────────────────────────────
      for (const value of BMS_CONSTANTS) {
        const item = new vscode.CompletionItem(
          value,
          vscode.CompletionItemKind.Constant
        );
        item.insertText = value;
        item.detail = 'BMS Constant';
        items.push(item);
      }

      // ───────────────────────────────
      // COLORS → values
      // ───────────────────────────────
      for (const value of BMS_COLORS) {
        const item = new vscode.CompletionItem(
          value,
          vscode.CompletionItemKind.Constant
        );
        item.insertText = value;
        item.detail = 'BMS Color';
        items.push(item);
      }
  
      return items;
    }
  }
  