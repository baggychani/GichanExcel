import JSZip from "@progress/jszip-esm";
import {
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  TextDecoration,
  VerticalAlign,
  WrapStrategy,
  type IStyleData,
} from "@univerjs/core";

export const BUILT_IN_NUMBER_FORMATS: Record<number, string> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)",
  38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mmss.0",
  48: "##0.0E+0",
  49: "@",
};

const INDEXED_COLORS = [
  "#000000",
  "#FFFFFF",
  "#FF0000",
  "#00FF00",
  "#0000FF",
  "#FFFF00",
  "#FF00FF",
  "#00FFFF",
  "#000000",
  "#FFFFFF",
  "#FF0000",
  "#00FF00",
  "#0000FF",
  "#FFFF00",
  "#FF00FF",
  "#00FFFF",
  "#800000",
  "#008000",
  "#000080",
  "#808000",
  "#800080",
  "#008080",
  "#C0C0C0",
  "#808080",
  "#9999FF",
  "#993366",
  "#FFFFCC",
  "#CCFFFF",
  "#660066",
  "#FF8080",
  "#0066CC",
  "#CCCCFF",
  "#000080",
  "#FF00FF",
  "#FFFF00",
  "#00FFFF",
  "#800080",
  "#800000",
  "#008080",
  "#0000FF",
  "#00CCFF",
  "#CCFFFF",
  "#CCFFCC",
  "#FFFF99",
  "#99CCFF",
  "#FF99CC",
  "#CC99FF",
  "#FFCC99",
  "#3366FF",
  "#33CCCC",
  "#99CC00",
  "#FFCC00",
  "#FF9900",
  "#FF6600",
  "#666699",
  "#969696",
  "#003366",
  "#339966",
  "#003300",
  "#333300",
  "#993300",
  "#993366",
  "#333399",
  "#333333",
];

export type OpenXmlStyleInfo = {
  zip: JSZip;
  styles: IStyleData[];
  sheetPathsByName: Map<string, string>;
};

export type OpenXmlSheetInfo = {
  cellStylesByRef: Map<string, IStyleData>;
};

function childElements(element: Element, localName?: string): Element[] {
  return Array.from(element.children).filter(
    (child) => !localName || child.localName === localName,
  );
}

function firstChildElement(element: Element, localName: string): Element | null {
  return childElements(element, localName)[0] ?? null;
}

function parseXml(text: string): Document {
  return new DOMParser().parseFromString(text, "application/xml");
}

function normalizeZipPath(path: string): string {
  const parts: string[] = [];
  path
    .replace(/^\/+/, "")
    .split("/")
    .forEach((part) => {
      if (!part || part === ".") {
        return;
      }
      if (part === "..") {
        parts.pop();
        return;
      }
      parts.push(part);
    });
  return parts.join("/");
}

function normalizeWorkbookTarget(target: string): string {
  return normalizeZipPath(target.startsWith("/") ? target : `xl/${target}`);
}

function rgbFromArgb(value: string | null): string | null {
  if (!value) {
    return null;
  }

  const normalized = value.replace(/[^0-9a-f]/gi, "").toUpperCase();
  if (normalized.length < 6) {
    return null;
  }

  return `#${normalized.slice(-6)}`;
}

function tintChannel(channel: number, tint: number): number {
  const result = tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint;
  return Math.max(0, Math.min(255, Math.round(result)));
}

function applyTint(color: string, tint: number): string {
  const hex = color.replace("#", "");
  if (hex.length !== 6 || !Number.isFinite(tint) || tint === 0) {
    return color;
  }

  const red = tintChannel(parseInt(hex.slice(0, 2), 16), tint);
  const green = tintChannel(parseInt(hex.slice(2, 4), 16), tint);
  const blue = tintChannel(parseInt(hex.slice(4, 6), 16), tint);
  return `#${[red, green, blue]
    .map((channel) => channel.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase()}`;
}

function parseColor(element: Element | null, themeColors: string[]): string | null {
  if (!element) {
    return null;
  }

  const rgb = rgbFromArgb(element.getAttribute("rgb"));
  if (rgb) {
    return rgb;
  }

  const indexed = Number(element.getAttribute("indexed"));
  if (Number.isInteger(indexed) && INDEXED_COLORS[indexed]) {
    return INDEXED_COLORS[indexed];
  }

  const themeIndex = Number(element.getAttribute("theme"));
  if (Number.isInteger(themeIndex) && themeColors[themeIndex]) {
    const tint = Number(element.getAttribute("tint") ?? "0");
    return applyTint(themeColors[themeIndex], tint);
  }

  return null;
}

function parseThemeColors(themeXml: string | null): string[] {
  if (!themeXml) {
    return [];
  }

  const doc = parseXml(themeXml);
  const scheme = Array.from(doc.getElementsByTagName("*")).find(
    (element) => element.localName === "clrScheme",
  );
  if (!scheme) {
    return [];
  }

  const themeKeys = [
    "lt1",
    "dk1",
    "lt2",
    "dk2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hlink",
    "folHlink",
  ];

  return themeKeys.map((key) => {
    const colorNode = firstChildElement(scheme, key);
    const srgb = colorNode
      ? Array.from(colorNode.getElementsByTagName("*")).find(
          (element) => element.localName === "srgbClr",
        )
      : null;
    const system = colorNode
      ? Array.from(colorNode.getElementsByTagName("*")).find(
          (element) => element.localName === "sysClr",
        )
      : null;
    return (
      rgbFromArgb(srgb?.getAttribute("val") ?? null) ??
      rgbFromArgb(system?.getAttribute("lastClr") ?? null) ??
      ""
    );
  });
}

function mapBorderStyle(style: string | null): BorderStyleTypes | null {
  switch (style) {
    case "thin":
      return BorderStyleTypes.THIN;
    case "hair":
      return BorderStyleTypes.HAIR;
    case "dotted":
      return BorderStyleTypes.DOTTED;
    case "dashed":
      return BorderStyleTypes.DASHED;
    case "dashDot":
      return BorderStyleTypes.DASH_DOT;
    case "dashDotDot":
      return BorderStyleTypes.DASH_DOT_DOT;
    case "double":
      return BorderStyleTypes.DOUBLE;
    case "medium":
      return BorderStyleTypes.MEDIUM;
    case "mediumDashed":
      return BorderStyleTypes.MEDIUM_DASHED;
    case "mediumDashDot":
      return BorderStyleTypes.MEDIUM_DASH_DOT;
    case "mediumDashDotDot":
      return BorderStyleTypes.MEDIUM_DASH_DOT_DOT;
    case "slantDashDot":
      return BorderStyleTypes.SLANT_DASH_DOT;
    case "thick":
      return BorderStyleTypes.THICK;
    default:
      return null;
  }
}

function mapHorizontalAlign(value: string | null): HorizontalAlign | undefined {
  switch (value) {
    case "left":
      return HorizontalAlign.LEFT;
    case "center":
    case "centerContinuous":
      return HorizontalAlign.CENTER;
    case "right":
      return HorizontalAlign.RIGHT;
    case "justify":
      return HorizontalAlign.JUSTIFIED;
    case "distributed":
      return HorizontalAlign.DISTRIBUTED;
    default:
      return undefined;
  }
}

function mapVerticalAlign(value: string | null): VerticalAlign | undefined {
  switch (value) {
    case "top":
      return VerticalAlign.TOP;
    case "center":
      return VerticalAlign.MIDDLE;
    case "bottom":
      return VerticalAlign.BOTTOM;
    default:
      return undefined;
  }
}

export function isStyleEmpty(style: IStyleData): boolean {
  return Object.keys(style).length === 0;
}

function parseOpenXmlStyles(stylesXml: string | null, themeColors: string[]): IStyleData[] {
  if (!stylesXml) {
    return [];
  }

  const doc = parseXml(stylesXml);
  const styleSheet = doc.documentElement;
  const customNumberFormats = new Map<number, string>();
  childElements(firstChildElement(styleSheet, "numFmts") ?? styleSheet, "numFmt").forEach(
    (numFmt) => {
      const id = Number(numFmt.getAttribute("numFmtId"));
      const formatCode = numFmt.getAttribute("formatCode");
      if (Number.isInteger(id) && formatCode) {
        customNumberFormats.set(id, formatCode);
      }
    },
  );
  const fonts = childElements(firstChildElement(styleSheet, "fonts") ?? styleSheet, "font");
  const fills = childElements(firstChildElement(styleSheet, "fills") ?? styleSheet, "fill");
  const borders = childElements(firstChildElement(styleSheet, "borders") ?? styleSheet, "border");
  const cellXfs = childElements(firstChildElement(styleSheet, "cellXfs") ?? styleSheet, "xf");

  const fontStyles = fonts.map((font): IStyleData => {
    const style: IStyleData = {};
    if (firstChildElement(font, "b")) {
      style.bl = BooleanNumber.TRUE;
    }
    if (firstChildElement(font, "i")) {
      style.it = BooleanNumber.TRUE;
    }
    if (firstChildElement(font, "u")) {
      style.ul = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE };
    }
    if (firstChildElement(font, "strike")) {
      style.st = { s: BooleanNumber.TRUE, t: TextDecoration.SINGLE };
    }

    const size = Number(firstChildElement(font, "sz")?.getAttribute("val"));
    if (Number.isFinite(size) && size > 0) {
      style.fs = size;
    }

    const color = parseColor(firstChildElement(font, "color"), themeColors);
    if (color) {
      style.cl = { rgb: color };
    }

    return style;
  });

  const fillStyles = fills.map((fill): IStyleData => {
    const pattern = firstChildElement(fill, "patternFill");
    const patternType = pattern?.getAttribute("patternType");
    if (!pattern || !patternType || patternType === "none") {
      return {};
    }

    const color = parseColor(
      firstChildElement(pattern, "fgColor") ?? firstChildElement(pattern, "bgColor"),
      themeColors,
    );
    return color ? { bg: { rgb: color } } : {};
  });

  const borderStyles = borders.map((border): IStyleData => {
    const bd: NonNullable<IStyleData["bd"]> = {};
    const directions = [
      ["top", "t"],
      ["right", "r"],
      ["bottom", "b"],
      ["left", "l"],
    ] as const;

    directions.forEach(([xmlName, univerName]) => {
      const borderElement = firstChildElement(border, xmlName);
      const mapped = mapBorderStyle(borderElement?.getAttribute("style") ?? null);
      if (!borderElement || mapped === null) {
        return;
      }

      const color = parseColor(firstChildElement(borderElement, "color"), themeColors);
      bd[univerName] = {
        s: mapped,
        cl: { rgb: color ?? "#000000" },
      };
    });

    return Object.keys(bd).length ? { bd } : {};
  });

  return cellXfs.map((xf) => {
    const style: IStyleData = {};
    const fontId = Number(xf.getAttribute("fontId"));
    const fillId = Number(xf.getAttribute("fillId"));
    const borderId = Number(xf.getAttribute("borderId"));
    const numFmtId = Number(xf.getAttribute("numFmtId"));

    if (fontStyles[fontId] && !isStyleEmpty(fontStyles[fontId])) {
      Object.assign(style, fontStyles[fontId]);
    }
    if (fillStyles[fillId] && !isStyleEmpty(fillStyles[fillId])) {
      Object.assign(style, fillStyles[fillId]);
    }
    if (borderStyles[borderId] && !isStyleEmpty(borderStyles[borderId])) {
      Object.assign(style, borderStyles[borderId]);
    }

    const numberFormat = customNumberFormats.get(numFmtId) ?? BUILT_IN_NUMBER_FORMATS[numFmtId];
    if (numberFormat && numberFormat !== "General") {
      style.n = { pattern: numberFormat };
    }

    const alignment = firstChildElement(xf, "alignment");
    if (alignment) {
      const horizontal = mapHorizontalAlign(alignment.getAttribute("horizontal"));
      const vertical = mapVerticalAlign(alignment.getAttribute("vertical"));
      if (horizontal) {
        style.ht = horizontal;
      }
      if (vertical) {
        style.vt = vertical;
      }
      if (alignment.getAttribute("wrapText") === "1") {
        style.tb = WrapStrategy.WRAP;
      }
    }

    return style;
  });
}

export async function readOpenXmlStyleInfo(buffer: ArrayBuffer): Promise<OpenXmlStyleInfo | null> {
  try {
    const zip = await new JSZip().loadAsync(buffer);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("text");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
    if (!workbookXml || !relsXml) {
      return null;
    }

    const themeXml = (await zip.file("xl/theme/theme1.xml")?.async("text")) ?? null;
    const themeColors = parseThemeColors(themeXml);
    const stylesXml = (await zip.file("xl/styles.xml")?.async("text")) ?? null;
    const styles = parseOpenXmlStyles(stylesXml, themeColors);
    const workbookDoc = parseXml(workbookXml);
    const relsDoc = parseXml(relsXml);
    const targetsById = new Map<string, string>();
    const sheetPathsByName = new Map<string, string>();

    Array.from(relsDoc.getElementsByTagName("*")).forEach((relationship) => {
      if (relationship.localName !== "Relationship") {
        return;
      }

      const id = relationship.getAttribute("Id");
      const target = relationship.getAttribute("Target");
      if (id && target) {
        targetsById.set(id, normalizeWorkbookTarget(target));
      }
    });

    Array.from(workbookDoc.getElementsByTagName("*")).forEach((sheet) => {
      if (sheet.localName !== "sheet") {
        return;
      }

      const name = sheet.getAttribute("name");
      const relId =
        sheet.getAttribute("r:id") ??
        sheet.getAttributeNS(
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "id",
        );
      const path = relId ? targetsById.get(relId) : null;
      if (name && path) {
        sheetPathsByName.set(name, path);
      }
    });

    return { zip, styles, sheetPathsByName };
  } catch (error) {
    console.warn("Failed to read XLSX workbook style metadata.", error);
    return null;
  }
}

export async function readOpenXmlSheetInfo(
  styleInfo: OpenXmlStyleInfo | null,
  sheetName: string,
): Promise<OpenXmlSheetInfo | null> {
  if (!styleInfo) {
    return null;
  }

  try {
    const sheetPath = styleInfo.sheetPathsByName.get(sheetName);
    if (!sheetPath) {
      return null;
    }

    const sheetXml = await styleInfo.zip.file(sheetPath)?.async("text");
    if (!sheetXml) {
      return null;
    }

    const doc = parseXml(sheetXml);
    const cellStylesByRef = new Map<string, IStyleData>();
    Array.from(doc.getElementsByTagName("*")).forEach((cell) => {
      if (cell.localName !== "c") {
        return;
      }

      const ref = cell.getAttribute("r");
      const styleIndex = Number(cell.getAttribute("s"));
      const style = styleInfo.styles[styleIndex];
      if (ref && style && !isStyleEmpty(style)) {
        cellStylesByRef.set(ref, style);
      }
    });

    return { cellStylesByRef };
  } catch (error) {
    console.warn(`Failed to read XLSX sheet style metadata for "${sheetName}".`, error);
    return null;
  }
}

