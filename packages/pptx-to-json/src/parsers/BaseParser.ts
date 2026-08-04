/**
 * Base parser utilities for PowerPoint component parsing
 */

import {
  emuToPixels,
  emuToPoints,
  validatePixelRange,
} from "../utils/constants.js";
import { ComponentStyle, XMLNode, TextRun } from "../types/index.js";

// Worker-compatible utility functions
export function isBufferLike(obj: any): boolean {
  return (
    obj &&
    (obj instanceof Uint8Array ||
      obj instanceof ArrayBuffer ||
      (typeof Buffer !== "undefined" &&
        Buffer.isBuffer &&
        Buffer.isBuffer(obj)))
  );
}

export function bufferFrom(data: any): Uint8Array | Buffer {
  if (typeof Buffer !== "undefined" && Buffer.from) {
    return Buffer.from(data);
  }
  return new Uint8Array(data);
}

// Font information interface
export interface FontInfo extends ComponentStyle {
  family: string;
  size: number;
  weight: string;
  style: string;
  decoration: string;
  color: string;
  isBold: boolean;
  isItalic: boolean;
  isUnderline: boolean;
  isSuperscript: boolean;
  isSubscript: boolean;
  isStrikethrough: boolean;
}

// Transform information interface
export interface TransformInfo {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
}

export class BaseParser {
  // Static theme colors - set once per document parsing session
  private static currentThemeColors: Record<string, string> | undefined;

  /**
   * Set the current theme colors for this parsing session
   * @param themeColors - Theme colors from extracted theme data
   */
  static setThemeColors(themeColors: Record<string, string> | undefined) {
    this.currentThemeColors = themeColors;
  }

  /**
   * Clear theme colors (call after parsing session)
   */
  static clearThemeColors() {
    this.currentThemeColors = undefined;
  }

  /**
   * Generate a clean, unique component ID
   * @param type - Component type (text, image, shape, etc.)
   * @param index - Component index within its type
   * @param originalName - Original PowerPoint name (optional, for debugging)
   * @returns Clean component ID suitable for CSS selectors
   */
  static generateComponentId(type: string, index: number, originalName?: string): string {
    // Create a base ID that's safe for CSS selectors
    const baseId = `${type}-${index}`;
    
    // If we have an original name, create a short hash for uniqueness
    if (originalName && originalName.trim()) {
      // Simple hash function to create a short identifier
      let hash = 0;
      for (let i = 0; i < originalName.length; i++) {
        const char = originalName.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash; // Convert to 32-bit integer
      }
      // Convert to positive hex string
      const hashStr = Math.abs(hash).toString(16).substring(0, 6);
      return `${baseId}-${hashStr}`;
    }
    
    return baseId;
  }

  /**
   * Convert PowerPoint font size units to points
   * PowerPoint uses hundreds of a point (1 point = 100 units)
   * @param sz - PowerPoint font size
   * @returns points
   */
  static fontSizeToPoints(sz: number | null | undefined): number {
    if (!sz || typeof sz !== "number") return 12; // Default font size
    return sz / 100;
  }

  /**
   * Parse color from PowerPoint color definition
   * @param colorDef - Color definition from PowerPoint XML
   * @returns hex color (#rrggbb)
   */
  static parseColor(colorDef: XMLNode | null | undefined): string {
    if (!colorDef) return "#000000";

    // Direct RGB color
    if (colorDef["srgbClr"]) {
      const val = colorDef["srgbClr"].$val;
      if (val) return `#${val}`;
    }

    // System color
    if (colorDef["sysClr"]) {
      const lastClr = colorDef["sysClr"].$lastClr;
      if (lastClr) return `#${lastClr}`;
    }

    // Scheme color (theme colors) - use static theme colors if available, fallback to defaults
    if (colorDef["schemeClr"]) {
      const val = colorDef["schemeClr"].$val;

      // Try to use static theme colors if set
      if (this.currentThemeColors && this.currentThemeColors[val]) {
        return this.currentThemeColors[val];
      }

      // Fallback to static defaults if no theme data available
      const defaultSchemeColors: Record<string, string> = {
        dk1: "#000000", // Dark 1
        lt1: "#FFFFFF", // Light 1
        dk2: "#44546A", // Dark 2
        lt2: "#E7E6E6", // Light 2
        tx1: "#000000", // Text 1 (same as dk1)
        tx2: "#44546A", // Text 2 (same as dk2)
        bg1: "#FFFFFF", // Background 1 (same as lt1)
        bg2: "#E7E6E6", // Background 2 (same as lt2)
        accent1: "#4472C4", // Accent 1
        accent2: "#E7686B", // Accent 2
        accent3: "#A5A5A5", // Accent 3
        accent4: "#FFC000", // Accent 4
        accent5: "#5B9BD5", // Accent 5
        accent6: "#70AD47", // Accent 6
        hlink: "#0563C1", // Hyperlink
        folHlink: "#954F72", // Followed hyperlink
      };
      return defaultSchemeColors[val] || "#000000";
    }

    return "#000000";
  }

  /**
   * Extract transform information (position, size, rotation)
   * @param xfrm - Transform object from PowerPoint
   * @returns transform data
   */
  static parseTransform(xfrm: XMLNode | import('../types/xml-nodes.js').TransformNode | null | undefined): TransformInfo {
    // Early return with defaults if not an object
    if (!xfrm || typeof xfrm !== 'object') {
      return { x: 0, y: 0, width: 0, height: 0, rotation: 0 };
    }

    // We purposely tolerate loosely shaped objects; perform presence checks guardedly

    const result: TransformInfo = {
      x: 0,
      y: 0,
      width: 0,
      height: 0,
      rotation: 0,
    };

    if (!xfrm) return result;

    // Position offset
    if (xfrm["off"]) {
      const off = xfrm["off"];
      if (off) {
        result.x = emuToPixels(parseInt(off.$x || 0));
        result.y = emuToPixels(parseInt(off.$y || 0));
      }
    }

    // Size extents
    if (xfrm["ext"]) {
      const ext = xfrm["ext"];
      if (ext) {
        result.width = emuToPixels(parseInt(ext.$cx || 0));
        result.height = emuToPixels(parseInt(ext.$cy || 0));
      }
    }

    // Rotation (in 60000ths of a degree)
    if (xfrm.$rot) {
      result.rotation = parseInt(xfrm.$rot) / 60000;
    }

    // Validate that we're returning pixel values, not EMU
    validatePixelRange(result.x, "transform x");
    validatePixelRange(result.y, "transform y");
    validatePixelRange(result.width, "transform width");
    validatePixelRange(result.height, "transform height");

    return result;
  }

  /**
   * Check if paragraph properties indicate bullet formatting
   * @param pPr - Paragraph properties from PowerPoint XML
   * @returns true if paragraph has bullet formatting
   */
  static hasBulletFormatting(pPr: XMLNode | null | undefined): boolean {
    if (!pPr) return false;

    // Check for bullet number properties
    if (pPr["buNum"] || pPr["buChar"]) {
      return true;
    }

    // Check for list properties
    if (pPr["lvl"] && pPr["lvl"] !== "0") {
      return true;
    }

    // Check for bullet fonts or autonum
    if (pPr["buAutoNum"] || pPr["buFont"]) {
      return true;
    }

    return false;
  }

  /**
   * Extract text content from PowerPoint text elements
   * @param textBody - Text body object from PowerPoint
   * @returns combined text content
   */
  static extractTextContent(textBody: XMLNode | null | undefined): string {
    if (!textBody || !textBody["p"]) return "";

    // Handle both array and single paragraph formats
    const paragraphs = Array.isArray(textBody["p"])
      ? textBody["p"]
      : [textBody["p"]];
    const textParts: string[] = [];

    paragraphs.forEach((paragraph, pIndex) => {
      let paragraphText = "";

      // Check for bullet formatting in paragraph properties
      const pPr = this.safeGet(paragraph, "pPr");
      const hasBullet = this.hasBulletFormatting(pPr);

      if (paragraph["r"]) {
        // Handle both array and single run formats
        const runs = Array.isArray(paragraph["r"])
          ? paragraph["r"]
          : [paragraph["r"]];
        runs.forEach((run) => {
          if (run["t"] !== undefined && run["t"] !== null) {
            // In fast-xml-parser, text is directly a string, not an array
            const text = run["t"];
            if (typeof text === "string") {
              paragraphText += text;
            } else if (typeof text === "number" || typeof text === "boolean") {
              // `parseTagValue` turns a run whose whole text is numeric into a NUMBER, so
              // a heading of "2026" is not a string. Dropping it leaves no trace at all.
              paragraphText += String(text);
            } else if (text._ && typeof text._ === "string") {
              paragraphText += text._;
            }
          }
        });
      }

      // Add bullet prefix if this paragraph has bullet formatting
      if (hasBullet && paragraphText.trim()) {
        paragraphText = "• " + paragraphText;
      }

      if (paragraphText) {
        textParts.push(paragraphText);
      }

      // Add paragraph break except for last paragraph
      if (pIndex < paragraphs.length - 1) {
        textParts.push("\n");
      }
    });

    return textParts.join("").trim();
  }

  /**
   * Extract font information from text run properties
   * @param rPr - Run properties from PowerPoint
   * @returns font information
   */
  static parseFont(rPr: XMLNode | null | undefined): FontInfo {
    const font: FontInfo = {
      family: "Arial",
      size: 12,
      weight: "normal",
      style: "normal",
      decoration: "none",
      color: "#000000",
      // Additional formatting properties
      isBold: false,
      isItalic: false,
      isUnderline: false,
      isSuperscript: false,
      isSubscript: false,
      isStrikethrough: false,
    };

    if (!rPr) return font;

    // Font family
    if (rPr["latin"]) {
      font.family = rPr["latin"].$typeface || font.family;
    }

    // Font size (in hundreds of a point)
    if (rPr.$sz) {
      font.size = this.fontSizeToPoints(parseInt(rPr.$sz));
    }

    // Bold
    if (rPr.$b === 1 || rPr.$b === "1") {
      font.weight = "bold";
      font.isBold = true;
    }

    // Italic
    if (rPr.$i === 1 || rPr.$i === "1") {
      font.style = "italic";
      font.isItalic = true;
    }

    // Underline
    if (rPr.$u && rPr.$u !== "none") {
      font.decoration = "underline";
      font.isUnderline = true;
    }

    // Strike-through
    if (rPr.$strike && rPr.$strike !== "noStrike") {
      font.isStrikethrough = true;
    }

    // Superscript/Subscript (baseline attribute)
    if (rPr.$baseline) {
      const baseline = parseInt(rPr.$baseline);
      if (baseline > 0) {
        font.isSuperscript = true;
      } else if (baseline < 0) {
        font.isSubscript = true;
      }
    }

    // Color
    if (rPr["solidFill"]) {
      font.color = this.parseColor(rPr["solidFill"]);
    }

    return font;
  }

  /**
   * Generate unique component ID
   * @param type - Component type
   * @param index - Component index
   * @returns unique ID
   */
  static generateId(type: string, index: number): string {
    return `${type}-${Date.now()}-${index}`;
  }

  /**
   * Fix spacing between text runs that may have been lost during PowerPoint parsing
   * @param textRuns - Array of text run objects
   * @returns Fixed text runs with proper spacing
   */
  static fixSpacingInTextRuns(textRuns: TextRun[]): TextRun[] {
    if (!textRuns || textRuns.length <= 1) return textRuns;

    const fixedRuns: TextRun[] = [];

    for (let i = 0; i < textRuns.length; i++) {
      const currentRun = textRuns[i];
      const nextRun = textRuns[i + 1];

      fixedRuns.push(currentRun);

      // Check if we need to add a space between this run and the next
      if (nextRun && this.shouldAddSpaceBetweenRuns(currentRun, nextRun)) {
        // Add a space as a separate text run
        fixedRuns.push({
          text: " ",
        });
      }
    }

    return fixedRuns;
  }

  /**
   * Determine if a space should be added between two text runs
   * @param run1 - First text run
   * @param run2 - Second text run
   * @returns true if space should be added
   */
  static shouldAddSpaceBetweenRuns(run1: TextRun, run2: TextRun): boolean {
    if (!run1?.text || !run2?.text) return false;

    const text1 = run1.text.trim();
    const text2 = run2.text.trim();

    // Don't add space if either text is empty
    if (!text1 || !text2) return false;

    // Don't add space if first text ends with whitespace or punctuation
    if (/[\s\.,!?;:]$/.test(run1.text)) return false;

    // Don't add space if second text starts with whitespace or punctuation
    if (/^[\s\.,!?;:]/.test(run2.text)) return false;

    // Add space between word-like runs
    return /\w$/.test(text1) && /^\w/.test(text2);
  }

  /**
   * Safe attribute access
   * @param obj - Object to access
   * @param path - Dot-notation path
   * @param defaultValue - Default value if path doesn't exist
   * @returns value or default
   */
  /**
   * Type-guard: check if a value looks like an XML node object
   */
  static isXMLNode(v: any): v is XMLNode {
    return v !== null && typeof v === 'object' && !Array.isArray(v);
  }

  /**
   * Type-guard: simple string detector (handles fast-xml-parser text nodes)
   */
  static isString(v: any): v is string {
    return typeof v === 'string' || (v && typeof v._ === 'string') || (v && typeof v.$val === 'string');
  }

  /**
   * Coerce a value to a safe string. Handles several XML node shapes used by the
   * parser (direct string, `{ _: string }`, `{ $val: string }`). Falls back to
   * the provided `fallback` when conversion is not possible.
   */
  static asString(v: any, fallback: string = ''): string {
    if (typeof v === 'string') return v;
    if (v && typeof v._ === 'string') return v._;
    if (v && typeof v.$val === 'string') return v.$val;
    if (v === undefined || v === null) return fallback;
    try {
      return String(v);
    } catch (_err) {
      return fallback;
    }
  }

  /**
   * Coerce a value to a safe integer number. Handles numeric strings and
   * `{ $val: string }` wrappers. Returns `fallback` when conversion fails.
   */
  static asNumber(v: any, fallback: number = 0): number {
    if (typeof v === 'number' && !Number.isNaN(v)) return v;
    if (typeof v === 'string' && v.trim() !== '') {
      const parsed = parseInt(v, 10);
      if (!Number.isNaN(parsed)) return parsed;
    }
    if (v && typeof v.$val === 'string') {
      const parsed = parseInt(v.$val, 10);
      if (!Number.isNaN(parsed)) return parsed;
    }
    return fallback;
  }

  /**
   * Safe attribute access
   * @param obj - Object to access
   * @param path - Dot-notation path
   * @param defaultValue - Default value if path doesn't exist
   * @returns value or default
   */
  static safeGet(obj: any, path: string, defaultValue?: any): any {
    try {
      const val = path.split('.').reduce((current: any, key: string) => current?.[key], obj);
      if (val === undefined || val === null) {
        return defaultValue !== undefined ? defaultValue : null;
      }
      return val;
    } catch (error) {
      return defaultValue !== undefined ? defaultValue : null;
    }
  }

  /**
   * Typed helper: get an XML node at the given path or return `defaultValue`.
   */
  static getNode(obj: any, path: string, defaultValue: XMLNode | null = null): XMLNode | null {
    const val = this.safeGet(obj, path, undefined);
    if (val === undefined || val === null) return defaultValue;
    if (this.isXMLNode(val)) return val as XMLNode;
    return defaultValue;
  }

  /**
   * Typed helper: get a string at the given path and coerce it safely.
   */
  static getString(obj: any, path: string, fallback: string = ''): string {
    const val = this.safeGet(obj, path, undefined);
    return this.asString(val, fallback);
  }

  /**
   * Typed helper: get a number at the given path and coerce it safely.
   */
  static getNumber(obj: any, path: string, fallback: number = 0): number {
    const val = this.safeGet(obj, path, undefined);
    return this.asNumber(val, fallback);
  }

  /**
   * Typed helper: get an array at the given path. If the value is a single object,
   * it will be wrapped in an array. Returns `fallback` when missing.
   */
  static getArray(obj: any, path: string, fallback: any[] = []): any[] {
    const val = this.safeGet(obj, path, undefined);
    if (val === undefined || val === null) return fallback;
    return Array.isArray(val) ? val : [val];
  }

  /**
   * Typed helper: get a boolean-like value at the given path and coerce it.
   */
  static getBoolean(obj: any, path: string, fallback: boolean = false): boolean {
    const val = this.safeGet(obj, path, undefined);
    if (val === undefined || val === null) return fallback;
    if (typeof val === 'boolean') return val;
    if (typeof val === 'number') return val === 1;
    if (typeof val === 'string') {
      const s = val.trim().toLowerCase();
      if (s === '1' || s === 'true' || s === 'yes') return true;
      if (s === '0' || s === 'false' || s === 'no') return false;
      return fallback;
    }
    if (val && typeof val.$val === 'string') {
      const s = val.$val.trim().toLowerCase();
      if (s === '1' || s === 'true' || s === 'yes') return true;
      if (s === '0' || s === 'false' || s === 'no') return false;
    }
    return fallback;
  }

  /**
   * Check if a color is white or near-white (should be treated as transparent background)
   * @param color - Hex color string
   * @returns true if color should be treated as no background
   */
  static isWhiteOrNearWhite(color: string | undefined): boolean {
    if (!color || color === 'transparent') return true;

    // Common white/near-white colors that should be treated as no background
    const whiteColors = [
      '#FFFFFF', // Pure white (bg1, lt1)
      '#E7E6E6', // Light gray (bg2, lt2) - appears as light violet
      '#FEFEFE', // Near white
      '#F8F8F8', // Very light gray
      '#F5F5F5', // Another light gray
      '#EEEEEE', // Light gray variant
      '#F0F0F0', // Another light gray variant
    ];

    const upperColor = color.toUpperCase();
    return whiteColors.includes(upperColor);
  }
}
