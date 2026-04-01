import * as vscode from 'vscode';
import { parseBmsFields } from './BmsDocument';

export const diagnosticCollection = vscode.languages.createDiagnosticCollection('bms');

function labelRules(diagnostics: vscode.Diagnostic[], line: number, text: string) {
    // ───────────────────────────────
    // RULE 1: Label BMS inválido
    // ───────────────────────────────

    // Detecta macro BMS
    const macroMatch = text.match(/\b(DFHMSD|DFHMDI|DFHMDF)\b/);

    if (macroMatch) {
        const macroIndex = macroMatch.index;

        // Tudo antes da macro
        const beforeMacro = text.slice(0, macroIndex);

        // Deve haver algo nas colunas 0–7 (label ou espaços)
        const labelArea = beforeMacro.slice(0, 7);
        const afterLabelArea = beforeMacro.slice(7);

        const labelMatch = labelArea.match(/^([A-Za-z][A-Za-z0-9_-]*)/);

        if (labelArea.trim().length === 0)
            return; // Sem label, ignora

        // ❌ Espaço antes do label (não começa na coluna 0)
        if (/^\s+/.test(labelArea)) {
            const range = new vscode.Range(line, 0, line, labelArea.length);

            diagnostics.push(
                new vscode.Diagnostic(
                    range,
                    'O label BMS deve iniciar na coluna 0',
                    vscode.DiagnosticSeverity.Error
                )
            );
        }

        // ❌ Label maior que 7 caracteres
        if (labelMatch && labelMatch[1].length > 7) {
            const range = new vscode.Range(
                line,
                0,
                line,
                labelMatch[1].length
            );

            diagnostics.push(
                new vscode.Diagnostic(
                    range,
                    'O label BMS deve ter no máximo 7 caracteres (colunas 0–7)',
                    vscode.DiagnosticSeverity.Error
                )
            );
        }

        // ❌ Texto fora da área do label antes da macro
        if (afterLabelArea.trim().length > 0) {
            const start = 7;
            const end = macroIndex || 0;

            const range = new vscode.Range(line, start, line, end);

            diagnostics.push(
                new vscode.Diagnostic(
                    range,
                    'Não é permitido texto entre o label e a macro',
                    vscode.DiagnosticSeverity.Error
                )
            );
        }
    } else if (text.trim().length > 0 && !text.trim().startsWith('*')) {
        const labelArea = (text + " ".repeat(7)).slice(0, 7);

        if (/\S/.test(labelArea)) {
            const range = new vscode.Range(line, 0, line, labelArea.length);

            diagnostics.push(
                new vscode.Diagnostic(
                    range,
                    'Linhas sem macro não devem conter texto na área do label (colunas 0–7)',
                    vscode.DiagnosticSeverity.Error
                )
            );
        }
    }
}

export function validateBmsDocument(document: vscode.TextDocument) {
    if (document.languageId !== 'bms') {
        return;
    }

    const diagnostics: vscode.Diagnostic[] = [];

    for (let line = 0; line < document.lineCount; line++) {
        const text = document.lineAt(line).text;

        labelRules(diagnostics, line, text);

        // ───────────────────────────────
        // RULE 2: Continuation character rules
        // ───────────────────────────────

        // Matches: comma + spaces + single non-space char at end of line
        const continuationMatch = text.match(/,( *)?(\S)$/);

        if (continuationMatch) {
            const continuationChar = continuationMatch[2];
            const continuationIndex = text.length - 1;

            // Must be exactly column 72 (index 71)
            if (continuationIndex !== 71) {
                const range = new vscode.Range(
                    line,
                    continuationIndex,
                    line,
                    continuationIndex + 1
                );

                diagnostics.push(
                    new vscode.Diagnostic(
                        range,
                        'O caractere de continuação deve estar na coluna 72',
                        vscode.DiagnosticSeverity.Error
                    )
                );
            }

            // Must be X
            if (continuationChar !== 'X' && continuationChar !== '*') {
                const range = new vscode.Range(
                    line,
                    continuationIndex,
                    line,
                    continuationIndex + 1
                );

                diagnostics.push(
                    new vscode.Diagnostic(
                        range,
                        'Caractere de continuação inválido (esperado X ou *)',
                        vscode.DiagnosticSeverity.Error
                    )
                );
            }
        }

        // ───────────────────────────────
        // RULE 3: Comentário deve iniciar na coluna 0
        // ───────────────────────────────
        if (/^\s+\*/.test(text)) {
            const firstStarIndex = text.indexOf('*');

            const range = new vscode.Range(
                line,
                0,
                line,
                firstStarIndex + 1
            );

            diagnostics.push(
                new vscode.Diagnostic(
                    range,
                    'Comentário deve iniciar na coluna 0 (sem espaços antes do *)',
                    vscode.DiagnosticSeverity.Error
                )
            );
        }
    }

    try {
        const parsedFields = parseBmsFields(document.getText());
        const duplicateGroups = new Map<string, typeof parsedFields>();

        for (const field of parsedFields) {
            if (field.askip || field.isArray || !field.logicalId) {
                continue;
            }

            const key = field.logicalId.toUpperCase();
            const existing = duplicateGroups.get(key) ?? [];
            existing.push(field);
            duplicateGroups.set(key, existing);
        }

        duplicateGroups.forEach((group, logicalId) => {
            if (group.length < 2) {
                return;
            }

            group.forEach(field => {
                diagnostics.push(
                    new vscode.Diagnostic(
                        new vscode.Range(field.idLine, field.idStart, field.idLine, field.idEnd),
                        `Duplicate variable ID "${logicalId}". Each non-ASKIP variable ID must be unique.`,
                        vscode.DiagnosticSeverity.Error
                    )
                );
            });
        });

        const icFields = parsedFields.filter(field => field.ic);
        if (icFields.length > 1) {
            const summary = icFields
                .map(field => field.logicalId || `POS=(${field.row + 1},${field.col + 1})`)
                .join(', ');
            const effectiveField = icFields[icFields.length - 1];
            const message = `Multiple IC fields found: ${summary}. The last one ("${effectiveField.logicalId || `POS=(${effectiveField.row + 1},${effectiveField.col + 1})`}") is where the cursor starts.`;

            icFields.forEach(field => {
                const lineLen = field.startLine < document.lineCount
                    ? document.lineAt(field.startLine).text.length
                    : 1;
                diagnostics.push(
                    new vscode.Diagnostic(
                        new vscode.Range(field.startLine, 0, field.startLine, Math.max(1, lineLen)),
                        message,
                        vscode.DiagnosticSeverity.Warning
                    )
                );
            });
        } else if (icFields.length === 0) {
            const firstLineLen = document.lineCount > 0 ? document.lineAt(0).text.length : 0;
            diagnostics.push(
                new vscode.Diagnostic(
                    new vscode.Range(0, 0, 0, Math.max(1, firstLineLen)),
                    'No IC field is set. The default cursor location is the first position of the screen.',
                    vscode.DiagnosticSeverity.Warning
                )
            );
        }
    } catch (_err) {
        // Parsing errors must not prevent other diagnostics from being reported
    }

    diagnosticCollection.set(document.uri, diagnostics);
}
