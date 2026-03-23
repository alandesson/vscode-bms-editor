// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import { BmsCompletionProvider } from './BmsCompletionProvider';
import { validateBmsDocument, diagnosticCollection } from './BmsValidationProvider';
import { openBmsRenderer } from './BmsRenderer';

export function activate(context: vscode.ExtensionContext) {

	console.log('"bms-renderer" is now active!');

	// --------------------------------------------------
	// Setting Up the Completion Provider for BMS files
	// --------------------------------------------------
	context.subscriptions.push(
		vscode.languages.registerCompletionItemProvider(
			{ language: 'bms' },
			new BmsCompletionProvider(),
			...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')
		)
	);

	// -------------------------------------
	// Setting Up Validation for BMS files
	// -------------------------------------
	context.subscriptions.push(diagnosticCollection);

	// Initial validation
	if (vscode.window.activeTextEditor) {
		validateBmsDocument(vscode.window.activeTextEditor.document);
	}

	// Re-validate on change
	context.subscriptions.push(
		vscode.workspace.onDidChangeTextDocument(e =>
			validateBmsDocument(e.document)
		)
	);

	// Re-validate on open
	context.subscriptions.push(
		vscode.workspace.onDidOpenTextDocument(validateBmsDocument)
	);

	// --------------------------
	// Setting Up Renderer
	// --------------------------
	const renderCommand = vscode.commands.registerCommand(
		'bms-renderer.render',
		async (uri?: vscode.Uri) => {
			let document: vscode.TextDocument;
			if (uri) {
				document = await vscode.workspace.openTextDocument(uri);
			} else {
				const editor = vscode.window.activeTextEditor;
				if (!editor || editor.document.languageId !== 'bms') {
					return;
				}
				document = editor.document;
			}
			openBmsRenderer(context, document);
		}
	);

	context.subscriptions.push(renderCommand);
}

// This method is called when your extension is deactivated
export function deactivate() { }
