// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import { BmsCompletionProvider } from './BmsCompletionProvider';
import { validateBmsDocument, diagnosticCollection } from './BmsValidationProvider';
import { openBmsRenderer } from './BmsRenderer';

// Known IBM / Broadcom mainframe extensions that provide competing BMS language support.
const CONFLICTING_EXTENSION_IDS = [
	'IBM.zopeneditor',          // IBM Z Open Editor
	'broadcomMFD.cobol-language-support',
	'BroadcomMFD.broadcom-cobol-language-support',
	'Rocket.lsp-cobol-support',
];

function detectConflict(): vscode.Extension<unknown> | undefined {
	return CONFLICTING_EXTENSION_IDS
		.map(id => vscode.extensions.getExtension(id))
		.find(ext => ext !== undefined);
}

async function warnOnConflict(context: vscode.ExtensionContext): Promise<void> {
	const suppress = vscode.workspace
		.getConfiguration('bms-renderer')
		.get<boolean>('suppressConflictWarning', false);
	if (suppress) { return; }

	const conflict = detectConflict();
	if (!conflict) { return; }

	const FIX    = 'Fix file associations';
	const IGNORE = "Don't show again";

	const choice = await vscode.window.showWarningMessage(
		`BMS Renderer: "${conflict.packageJSON.displayName}" is also installed and may override ` +
		`BMS syntax highlighting and the renderer toolbar button. ` +
		`Adding \`"*.bms": "bms"\` to your \`files.associations\` setting restores correct behaviour.`,
		FIX,
		IGNORE
	);

	if (choice === FIX) {
		// Write the file association into the user's global settings.
		const filesConfig = vscode.workspace.getConfiguration('files');
		const current: Record<string, string> =
			filesConfig.inspect<Record<string, string>>('associations')?.globalValue ?? {};
		await filesConfig.update(
			'associations',
			{ ...current, '*.bms': 'bms' },
			vscode.ConfigurationTarget.Global
		);
		vscode.window.showInformationMessage(
			'BMS Renderer: added `"*.bms": "bms"` to your global files.associations. ' +
			'Reload the window for changes to take effect.',
			'Reload'
		).then(r => { if (r === 'Reload') { vscode.commands.executeCommand('workbench.action.reloadWindow'); } });
	} else if (choice === IGNORE) {
		await vscode.workspace
			.getConfiguration('bms-renderer')
			.update('suppressConflictWarning', true, vscode.ConfigurationTarget.Global);
	}
}

export function activate(context: vscode.ExtensionContext) {

	console.log('"bms-renderer" is now active!');

	// Check for conflicting mainframe extensions (IBM Z Open Editor, Broadcom, etc.)
	// and guide the user to fix file associations if needed.
	warnOnConflict(context);

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
				if (!editor) { return; }
				document = editor.document;
				// Accept the active document if it is the bms language ID OR if its
				// file extension is .bms / .bbms (covers the case where a conflicting
				// extension has claimed the language ID for .bms files).
				const ext = document.uri.fsPath.replace(/^.*\./, '').toLowerCase();
				if (document.languageId !== 'bms' && ext !== 'bms' && ext !== 'bbms') {
					return;
				}
			}
			openBmsRenderer(context, document);
		}
	);

	context.subscriptions.push(renderCommand);
}

// This method is called when your extension is deactivated
export function deactivate() { }
