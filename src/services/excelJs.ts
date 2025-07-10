import type ExcelJS from "exceljs";
import type { Cell } from "microsoft-graph/Cell";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";

export function appendRow(worksheet: ExcelJS.Worksheet, row: Partial<Cell>[]): void {
    const excelRow = worksheet.addRow(row.map((cell) => cell.value ?? ""));
    const rowIndex = excelRow.number; // ExcelJS row number (1-based)
    let colIndex = 1;
    for (const cell of row) {
        const excelCell = excelRow.getCell(colIndex);
        // Write format (number format)
        if (cell.format) {
            excelCell.numFmt = cell.format as string;
        }
        // Write merge (merge cells right/down)
        if (cell.merge && (cell.merge.right || cell.merge.down)) {
            const startCol = colIndex;
            const startRow = rowIndex;
            const endCol = startCol + (cell.merge.right ?? 0);
            const endRow = startRow + (cell.merge.down ?? 0);
            worksheet.mergeCells(startRow, startCol, endRow, endCol);
        }
        // Write alignment
        const mappedAlignment = mapAlignment(cell.alignment);
        if (mappedAlignment) {
            excelCell.alignment = mappedAlignment;
        }
        // Write borders
        const mappedBorders = mapBorders(cell.borders);
        if (mappedBorders) {
            excelCell.border = mappedBorders;
        }
        // Write fill
        const mappedFill = mapFill(cell.fill);
        if (mappedFill) {
            excelCell.fill = mappedFill;
        }
        // Write font
        const mappedFont = mapFont(cell.font);
        if (mappedFont) {
            excelCell.font = mappedFont;
        }
        colIndex++;
    }
}
function mapAlignment(alignment?: Cell["alignment"]): Partial<ExcelJS.Alignment> | undefined {
    if (!alignment) return undefined;
    const horizontalMap: Record<string, ExcelJS.Alignment["horizontal"]> = {
        Left: "left",
        Center: "center",
        Right: "right",
        Fill: "fill",
        Justify: "justify",
        CenterAcrossSelection: "centerContinuous",
        Distributed: "distributed",
    };
    const verticalMap: Record<string, ExcelJS.Alignment["vertical"]> = {
        Top: "top",
        Center: "middle",
        Bottom: "bottom",
        Justify: "justify",
        Distributed: "distributed",
    };
    let horizontal: ExcelJS.Alignment["horizontal"] | undefined;
    let vertical: ExcelJS.Alignment["vertical"] | undefined;
    if (alignment.horizontal) {
        if (!(alignment.horizontal in horizontalMap)) {
            throw new InvalidArgumentError(`Unsupported horizontal alignment: ${alignment.horizontal}`);
        }
        horizontal = horizontalMap[alignment.horizontal];
    }
    if (alignment.vertical) {
        if (!(alignment.vertical in verticalMap)) {
            throw new InvalidArgumentError(`Unsupported vertical alignment: ${alignment.vertical}`);
        }
        vertical = verticalMap[alignment.vertical];
    }
    const result: Partial<ExcelJS.Alignment> = {};
    if (horizontal) result.horizontal = horizontal;
    if (vertical) result.vertical = vertical;
    if (typeof alignment.wrapText === "boolean") result.wrapText = alignment.wrapText;
    return Object.keys(result).length > 0 ? result : undefined;
}
function mapBorders(borders?: Cell["borders"]): Partial<ExcelJS.Borders> | undefined {
    if (!borders) return undefined;
    const supported = ["edgeTop", "edgeBottom", "edgeLeft", "edgeRight"];
    for (const key of Object.keys(borders)) {
        if (!supported.includes(key)) {
            throw new InvalidArgumentError(`Unsupported border property: ${key}`);
        }
    }
    const mapBorder = (b: unknown): ExcelJS.Border | undefined => {
        if (!b) return undefined;
        if (typeof b !== "object") throw new InvalidArgumentError("Border must be an object");
        // You may want to add more mapping here
        return b as ExcelJS.Border;
    };
    const result: Partial<ExcelJS.Borders> = {};
    const top = mapBorder(borders.edgeTop);
    if (top) result.top = top;
    const bottom = mapBorder(borders.edgeBottom);
    if (bottom) result.bottom = bottom;
    const left = mapBorder(borders.edgeLeft);
    if (left) result.left = left;
    const right = mapBorder(borders.edgeRight);
    if (right) result.right = right;
    return Object.keys(result).length > 0 ? result : undefined;
}
function colorToARGB(color: string): string {
    // Only allow string color for now
    if (typeof color === "string") return color;
    throw new InvalidArgumentError("Unsupported color type for ExcelJS");
}
function mapFill(fill?: Cell["fill"]): ExcelJS.Fill | undefined {
    if (!fill || !fill.color) return undefined;
    if (typeof fill.color !== "string") throw new InvalidArgumentError("Unsupported fill color type for ExcelJS");
    return {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: colorToARGB(fill.color) },
    };
}
function mapFont(font?: Cell["font"]): Partial<ExcelJS.Font> | undefined {
    if (!font) return undefined;
    if (font.color && typeof font.color !== "string") {
        throw new InvalidArgumentError("Unsupported font color type for ExcelJS");
    }
    const result: Partial<ExcelJS.Font> = {};
    if (typeof font.name === "string") result.name = font.name;
    if (typeof font.size === "number") result.size = font.size;
    if (font.color) result.color = { argb: colorToARGB(font.color) };
    if (typeof font.bold === "boolean") result.bold = font.bold;
    if (typeof font.italic === "boolean") result.italic = font.italic;
    if (font.underline && font.underline !== "None") result.underline = true;
    return Object.keys(result).length > 0 ? result : undefined;
}


export async function xlsxToBuffer(workbook: ExcelJS.Workbook): Promise<Buffer> {
    return Buffer.from(await workbook.xlsx.writeBuffer());
}