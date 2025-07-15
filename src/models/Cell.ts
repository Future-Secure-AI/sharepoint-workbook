export type CellWrite = Partial<
	Omit<Cell, "text"> & {
		mergeDown: number | null;
		mergeRight: number | null;
	}
>;

export type Cell = {
	value: CellValue | null;
	text: string;
	format: Format | null;
	note: string | null;

	fontName: FontName | null;
	fontSize: number | null;
	fontFamily: number | null;
	fontColor: RgbaColor | null;
	fontBold: boolean | null;
	fontItalic: boolean | null;
	fontUnderline: FontUnderlineStyle | null;
	fontStrike: boolean | null;
	fontOutline: boolean | null;

	alignmentHorizontal: AlignmentHorizontal | null;
	alignmentVertical: AlignmentVertical | null;
	alignmentWrapText: boolean | null;
	alignmentShrinkToFit: boolean | null;
	alignmentIndent: number | null;
	alignmentTextRotation: AlignmentRotation | null;

	borderTopStyle: BorderStyle | null;
	borderTopColor: RgbaColor | null;
	borderLeftStyle: BorderStyle | null;
	borderLeftColor: RgbaColor | null;
	borderBottomStyle: BorderStyle | null;
	borderBottomColor: RgbaColor | null;
	borderRightStyle: BorderStyle | null;
	borderRightColor: RgbaColor | null;

	fillForegroundColor: RgbaColor | null;
	fillBackgroundColor: RgbaColor | null;

	protectionLocked: boolean | null;
	protectionHidden: boolean | null;
};

export type CellValue = string | number | boolean | Date;

export type Format = string;

export type FontName = string;
export type FontUnderlineStyle = "none" | "single" | "double" | "singleAccounting" | "doubleAccounting";

export type BorderStyle = "thin" | "dotted" | "hair" | "medium" | "double" | "thick" | "dashed" | "dashDot" | "dashDotDot" | "slantDashDot" | "mediumDashed" | "mediumDashDotDot" | "mediumDashDot";

export type AlignmentHorizontal = "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
export type AlignmentVertical = "top" | "middle" | "bottom" | "distributed" | "justify";
export type AlignmentRotation = number | "vertical";

/**
 * 8-character hex color string, e.g. "FF0000FF" for red with full opacity.
 */
export type RgbaColor = string;
