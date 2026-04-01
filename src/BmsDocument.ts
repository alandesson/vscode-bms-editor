export interface ParsedBmsField {
    label: string;
    logicalId: string;
    rawLogicalId: string;
    row: number;
    col: number;
    askip: boolean;
    ic: boolean;
    isArray: boolean;
    startLine: number;
    idLine: number;
    idStart: number;
    idEnd: number;
}

interface ParsedComment {
    leftId: string;
    rightId: string;
    isArray: boolean;
    line: number;
    rightStart: number;
    rightEnd: number;
}

function parseCommentLine(line: string, lineNumber: number): ParsedComment | null {
    if (!line.startsWith('* ')) {
        return null;
    }

    const leftMatch = line.slice(2).match(/^(\S+)/);
    if (!leftMatch) {
        return null;
    }

    const leftId = leftMatch[1];
    const rightSlice = line.length > 34 ? line.slice(34) : '';
    const rightMatch = rightSlice.match(/^(\S+)/);
    if (!rightMatch) {
        return null;
    }

    const rightId = rightMatch[1];
    const rightStart = 34;
    const rightEnd = rightStart + rightId.length;

    return {
        leftId,
        rightId,
        isArray: leftId === rightId,
        line: lineNumber,
        rightStart,
        rightEnd,
    };
}

function inferPrefix(fields: Array<{ label: string; comment: ParsedComment | null }>): string | null {
    // Detect prefix by splitting each non-array comment rightId at the first hyphen.
    // If the part before the hyphen is identical across all commented fields, that is the prefix.
    const commentedFields = fields.filter(f => f.comment && !f.comment.isArray);
    if (commentedFields.length < 1) {
        return null;
    }

    const firstParts = commentedFields.map(f => {
        const rightId = f.comment!.rightId;
        const hyphenIdx = rightId.indexOf('-');
        return hyphenIdx > 0 ? rightId.slice(0, hyphenIdx) : null;
    });

    if (!firstParts.every(p => p !== null)) {
        return null;
    }

    const firstUpper = firstParts[0]!.toUpperCase();
    if (!firstParts.every(p => p!.toUpperCase() === firstUpper)) {
        return null;
    }

    return firstParts[0];
}

function stripPrefix(value: string, prefix: string | null): string {
    if (!prefix) {
        return value;
    }

    const normalizedPrefix = `${prefix}-`;
    if (value.toUpperCase().startsWith(normalizedPrefix.toUpperCase())) {
        return value.slice(normalizedPrefix.length);
    }

    return value;
}

export function parseBmsFields(source: string): ParsedBmsField[] {
    const lines = source.split(/\r?\n/);
    const starts: Array<{ startLine: number; label: string }> = [];

    for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
        const line = lines[lineIndex];
        const labeledMatch = line.match(/^([A-Za-z][A-Za-z0-9-]{0,6})\s+DFHMDF\b/i);
        if (labeledMatch) {
            starts.push({ startLine: lineIndex, label: labeledMatch[1] });
            continue;
        }

        if (/^\s{2,}DFHMDF\b/i.test(line)) {
            starts.push({ startLine: lineIndex, label: '' });
        }
    }

    const stagedFields = starts.map((start, index) => {
        const endLine = index + 1 < starts.length ? starts[index + 1].startLine : lines.length;
        const block = lines.slice(start.startLine, endLine).join('\n');
        const comment = start.startLine > 0 ? parseCommentLine(lines[start.startLine - 1], start.startLine - 1) : null;

        const posMatch = block.match(/POS=\((\d+)\s*,\s*(\d+)\)/i);
        if (!posMatch) {
            return null;
        }

        // Use [^)] to allow newlines in the capture so multi-line ATTRB= continuation is handled.
        // Strip trailing continuation (,       X\n<spaces>) before splitting.
        const attrbMatch = block.match(/ATTRB=\(?([^)]*)\)?/i);
        const attrbStr = attrbMatch
            ? attrbMatch[1].replace(/,[ \t]*X[ \t]*\r?\n[ \t]*/g, ',')
            : '';
        const attrbList = attrbStr.split(',').map(item => item.trim().toUpperCase()).filter(Boolean);

        const rawLogicalId = comment && !comment.isArray ? comment.rightId : start.label;
        const idLine = comment && !comment.isArray ? comment.line : start.startLine;
        const idStart = comment && !comment.isArray ? comment.rightStart : 0;
        const idEnd = comment && !comment.isArray ? comment.rightEnd : start.label.length;

        return {
            startLine: start.startLine,
            label: start.label,
            comment,
            row: Number(posMatch[1]) - 1,
            col: Number(posMatch[2]) - 1,
            askip: attrbList.includes('ASKIP'),
            ic: attrbList.includes('IC'),
            rawLogicalId,
            idLine,
            idStart,
            idEnd,
        };
    }).filter((field): field is NonNullable<typeof field> => field !== null);

    const inferredPrefix = inferPrefix(stagedFields);

    return stagedFields.map(field => ({
        label: field.label,
        logicalId: stripPrefix(field.rawLogicalId, inferredPrefix),
        rawLogicalId: field.rawLogicalId,
        row: field.row,
        col: field.col,
        askip: field.askip,
        ic: field.ic,
        isArray: !!field.comment?.isArray,
        startLine: field.startLine,
        idLine: field.idLine,
        idStart: field.idStart,
        idEnd: field.idEnd,
    }));
}

export function findBmsFieldStartLine(source: string, row: number, col: number): number | undefined {
    return parseBmsFields(source).find(field => field.row === row && field.col === col)?.startLine;
}
