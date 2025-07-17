import AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";

export function applyCell(worksheet: AsposeCells.Worksheet, r: number, c: number, cellOrValue: CellValue | DeepPartial<Cell>) {
	const output = worksheet.cells.get(r, c);

	if (typeof cellOrValue === "object" && !(cellOrValue instanceof Date)) {
		const cell = cellOrValue as DeepPartial<Cell>;

		if (cell.value !== undefined) {
			if (typeof cell.value === "number") {
				output.putValue(cell.value);
			} else if (typeof cell.value === "boolean") {
				output.putValue(cell.value);
			} else if (typeof cell.value === "string") {
				output.putValue(cell.value);
			} else if (cell.value instanceof Date) {
				output.putValue(cell.value);
			} else {
				throw new InvalidArgumentError(`Unsupported cell value type: ${typeof cell.value}. Expected number, boolean, string, or Date.`);
			}
		}

		if (cell.formula !== undefined) {
			output.formula = cell.formula;
		}

		if (cell.style !== undefined) {
			const style = worksheet.workbook.createStyle();
			style.copy(output.getStyle());

			if (cell.style.font !== undefined) {
				if (cell.style.font.name !== undefined) style.font.setName(cell.style.font.name);
				if (cell.style.font.size !== undefined) style.font.size = cell.style.font.size;
				if (cell.style.font.bold !== undefined) style.font.isBold = cell.style.font.bold;
				if (cell.style.font.italic !== undefined) style.font.isItalic = cell.style.font.italic;
				if (cell.style.font.color !== undefined) style.font.color = parseColor(cell.style.font.color);
			}

			if (cell.style.backgroundColor !== undefined) {
				style.pattern = AsposeCells.BackgroundType.Solid;
				style.foregroundColor = parseColor(cell.style.backgroundColor);
			}

			if (cell.style.horizontalAlignment !== undefined) {
				const hAlignMap = {
					// TODO: More possible alignments
					left: AsposeCells.TextAlignmentType.Left,
					center: AsposeCells.TextAlignmentType.Center,
					right: AsposeCells.TextAlignmentType.Right,
				};
				style.horizontalAlignment = hAlignMap[cell.style.horizontalAlignment];
			}

			if (cell.style.verticalAlignment !== undefined) {
				const vAlignMap = {
					top: AsposeCells.TextAlignmentType.Top,
					middle: AsposeCells.TextAlignmentType.Center,
					bottom: AsposeCells.TextAlignmentType.Bottom,
				};
				style.verticalAlignment = vAlignMap[cell.style.verticalAlignment];
			}

			if (cell.style.borders !== undefined) {
				// TOOD: More possible borders
				const borderMap = {
					top: AsposeCells.BorderType.TopBorder,
					bottom: AsposeCells.BorderType.BottomBorder,
					left: AsposeCells.BorderType.LeftBorder,
					right: AsposeCells.BorderType.RightBorder,
				};
				const borderStyleMap = {
					thin: AsposeCells.CellBorderType.Thin,
					medium: AsposeCells.CellBorderType.Medium,
					thick: AsposeCells.CellBorderType.Thick,
					dashed: AsposeCells.CellBorderType.Dashed,
					dotted: AsposeCells.CellBorderType.Dotted,
					double: AsposeCells.CellBorderType.Double,
				};
				for (const side of ["top", "bottom", "left", "right"] as const) {
					const border = cell.style.borders[side];
					if (border && border.style !== undefined && border.color !== undefined) {
						const borderType = borderMap[side];
						const borderStyle = borderStyleMap[border.style as keyof typeof borderStyleMap];
						if (borderType !== undefined && borderStyle !== undefined) {
							style.setBorder(borderType, borderStyle, parseColor(border.color));
						}
					}
				}
			}

			if (cell.style.numberFormat !== undefined) {
				style.custom = cell.style.numberFormat;
			}

			if (cell.style.locked !== undefined) {
				style.isLocked = cell.style.locked;
			}

			if (cell.style.wrapText !== undefined) {
				style.isTextWrapped = cell.style.wrapText;
			}

			output.setStyle(style);
		}

		if (cell.merge !== undefined) {
			const mergedRegions = worksheet.cells.getMergedAreas();
			if (cell.merge === null) {
				// Unmerge this cell if it is part of a merged region
				for (let i = 0; i < mergedRegions.length; i++) {
					const region = mergedRegions[i];
					if (!region) continue;
					const regionStartRow = region.startRow;
					const regionStartCol = region.startColumn;
					const regionEndRow = region.endRow;
					const regionEndCol = region.endColumn;
					if (r >= regionStartRow && r <= regionEndRow && c >= regionStartCol && c <= regionEndCol) {
						const rowCount = regionEndRow - regionStartRow + 1;
						const colCount = regionEndCol - regionStartCol + 1;
						worksheet.cells.unMerge(regionStartRow, regionStartCol, rowCount, colCount);
					}
				}
			} else {
				// Merge cells according to the new CellMerge type: "up" | "left" | "up-left"
				let startRow = r;
				let startCol = c;
				let endRow = r;
				let endCol = c;

				if (cell.merge === "up") {
					startRow = r - 1;
				} else if (cell.merge === "left") {
					startCol = c - 1;
				} else if (cell.merge === "up-left") {
					startRow = r - 1;
					startCol = c - 1;
				}

				// Scan all merged regions to see if the merge neighbor (above/left) is already merged
				// If so, expand the merge area to include the entire region
				for (let i = 0; i < mergedRegions.length; i++) {
					const region = mergedRegions[i];
					if (!region) continue;
					const regionStartRow = region.startRow;
					const regionStartCol = region.startColumn;
					const regionEndRow = region.endRow;
					const regionEndCol = region.endColumn;

					// If the current cell or its merge neighbor is inside or adjacent to a region, expand the merge area
					if (
						(startRow >= regionStartRow && startRow <= regionEndRow && startCol >= regionStartCol && startCol <= regionEndCol) ||
						(endRow >= regionStartRow && endRow <= regionEndRow && endCol >= regionStartCol && endCol <= regionEndCol) ||
						(regionEndRow + 1 === startRow && regionStartCol === startCol && regionEndCol === endCol) || // above
						(regionEndCol + 1 === startCol && regionStartRow === startRow && regionEndRow === endRow) // left
					) {
						startRow = Math.min(startRow, regionStartRow);
						startCol = Math.min(startCol, regionStartCol);
						endRow = Math.max(endRow, regionEndRow);
						endCol = Math.max(endCol, regionEndCol);
					}
				}

				// Only merge if the startRow/startCol are valid
				if (startRow >= 0 && startCol >= 0) {
					// Remove any overlapping merged regions first by unmerging them
					for (let i = 0; i < mergedRegions.length; i++) {
						const region = mergedRegions[i];
						if (!region) continue;
						const regionStartRow = region.startRow;
						const regionStartCol = region.startColumn;
						const regionEndRow = region.endRow;
						const regionEndCol = region.endColumn;
						if (regionStartRow >= startRow && regionEndRow <= endRow && regionStartCol >= startCol && regionEndCol <= endCol) {
							const rowCount = regionEndRow - regionStartRow + 1;
							const colCount = regionEndCol - regionStartCol + 1;
							worksheet.cells.unMerge(regionStartRow, regionStartCol, rowCount, colCount);
						}
					}
					worksheet.cells.merge(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
				}
			}
		}

		if (cell.comment !== undefined) {
			const comment = worksheet.comments.get(worksheet.comments.add(r, c));
			comment.note = cell.comment;
		}
	} else {
		setCellValueAndFormula(output, cellOrValue);
	}
}

export function setCellValueAndFormula(cell: AsposeCells.Cell, value: CellValue) {
	if (typeof value === "number") {
		cell.putValue(value);
	} else if (typeof value === "boolean") {
		cell.putValue(value);
	} else if (typeof value === "string") {
		if (value.startsWith("=")) {
			cell.putValue("");
			cell.formula = value;
		} else {
			cell.putValue(value);
		}
	} else if (value instanceof Date) {
		cell.putValue(value); // Aspose will handle Date properly
	} else {
		throw new InvalidArgumentError(`Unsupported cell value type: ${typeof value}. Expected number, boolean, string, or Date.`);
	}
}

export function parseColor(hex: string): AsposeCells.Color {
	let color = hex.trim();
	if (color.startsWith("#")) {
		color = color.slice(1);
	}

	if (!/^[0-9a-fA-F]{6}$/.test(color)) {
		throw new InvalidArgumentError(`Invalid color string: '${hex}'. Expected a 6-digit hexadecimal string optionally prefixed with '#'.`);
	}

	const r = parseInt(color.slice(0, 2), 16);
	const g = parseInt(color.slice(2, 4), 16);
	const b = parseInt(color.slice(4, 6), 16);
	return new AsposeCells.Color(r, g, b);
}
