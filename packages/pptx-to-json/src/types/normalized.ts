/**
 * Normalized PowerPoint structure types
 *
 * These types represent the unified intermediate structure produced by
 * PowerPointNormalizer. They intentionally keep many fields as `any` until
 * downstream parsers are fully migrated. This prevents excessive churn while
 * still giving consumers a stable, discriminated union surface.
 */

export type NormalizedFormat = 'pptx' | 'clipboard';

export type NormalizedElement =
  | NormalizedTextElement
  | NormalizedShapeElement
  | NormalizedImageElement
  | NormalizedTableElement
  | NormalizedDiagramElement
  | NormalizedVideoElement
  | NormalizedConnectionElement;

/** Base shared fields for all normalized elements */
export interface NormalizedElementBase {
  type: 'text' | 'shape' | 'image' | 'table' | 'video' | 'connection' | 'diagram';
  /** Z-order index (higher renders on top) */
  zIndex: number;
  /** PowerPoint namespace origin ('p' = pptx slide, 'a' = clipboard drawing) */
  namespace: 'p' | 'a';
  /** Original element tag name (e.g. 'sp', 'pic', 'graphicFrame') */
  element: string;
  /** Original raw element data object */
  data: any;
  /** Flags for background sources */
  isLayoutElement?: boolean;
  isMasterElement?: boolean;
  isBackgroundElement?: boolean;
  /** Relationship id (images / videos) */
  relationshipId?: string;
}

export interface NormalizedTextElement extends NormalizedElementBase {
  type: 'text';
  spPr?: any;
  nvSpPr?: any;
  style?: any;
  /** Direct text body (pptx) */
  textBody?: any;
  /** Alternate text body (layout/master) */
  txBody?: any;
}

export interface NormalizedShapeElement extends NormalizedElementBase {
  type: 'shape';
  spPr?: any;
  nvSpPr?: any;
  style?: any;
  /** Empty text container or clipboard nested text */
  textBody?: any;
}

export interface NormalizedImageElement extends NormalizedElementBase {
  type: 'image';
  nvPicPr?: any;
  blipFill?: any;
  spPr?: any;
  style?: any;
}

export interface NormalizedVideoElement extends NormalizedElementBase {
  type: 'video';
  nvPicPr?: any;
  blipFill?: any;
  spPr?: any;
  style?: any;
  relationshipId?: string; // override to ensure present for videos
}

export interface NormalizedTableElement extends NormalizedElementBase {
  type: 'table';
  nvGraphicFramePr?: any;
  spPr?: any; // Positioning (xfrm)
  graphicData?: any; // Table XML contents
}

export interface NormalizedDiagramElement extends NormalizedElementBase {
  type: 'diagram';
  nvGraphicFramePr?: any;
  spPr?: any; // Positioning (xfrm)
  graphicData?: any; // Diagram XML contents
}

export interface NormalizedConnectionElement extends NormalizedElementBase {
  type: 'connection';
  nvCxnSpPr?: any; // Connection shape properties
  spPr?: any; // Line positioning and geometry
  startConnection?: any; // Connection start point
  endConnection?: any; // Connection end point
}

export interface NormalizedSlide {
  /** Source file path (ppt/slides/slide1.xml or clipboard/drawings/drawing1.xml) */
  slideFile: string;
  /** 1-based position in PRESENTATION order (p:sldIdLst); undefined for clipboard */
  slideNumber?: number;
  /**
   * The number in the slide's filename — creation order, NOT what the audience sees.
   * Kept because the rels graph is keyed on it: `ppt/slides/_rels/slide<n>.xml.rels`.
   */
  fileSlideNumber?: number;
  /** The slide's own relationship part, for resolving its media without arithmetic on slide numbers */
  slideRelsFile?: string;
  /** Text of the title placeholder, when the slide has one (27% of real slides do not) */
  title?: string;
  /** Whether the title placeholder named itself, or was named by the layout via its idx */
  titleSource?: 'placeholder' | 'layout';
  /** Speaker notes, when the slide has any beyond the auto-generated slide-number field */
  notes?: string;
  /** Notes part resolved through the rels graph — never paired by filename number */
  notesFile?: string | null;
  /** Original format: 'pptx' | 'clipboard' */
  format: NormalizedFormat;
  /** Legacy collections (pre-normalization parsing still references these) */
  shapes: any[];
  images: any[];
  text: any[];
  videos: any[];
  /** New ordered element list preserving z-order */
  elements: NormalizedElement[];
  /** Layout / master provenance */
  layoutFile?: string;
  masterFile?: string | null;
  layoutElementCount?: number;
  masterElementCount?: number;
  /** Raw trees for relationship lookups */
  rawSpTree?: any;
  rawCanvas?: any;
}

// Type guards (narrowers) for normalized elements
export function isTextElement(el: NormalizedElement): el is NormalizedTextElement { return el.type === 'text'; }
export function isShapeElement(el: NormalizedElement): el is NormalizedShapeElement { return el.type === 'shape'; }
export function isImageElement(el: NormalizedElement): el is NormalizedImageElement { return el.type === 'image'; }
export function isTableElement(el: NormalizedElement): el is NormalizedTableElement { return el.type === 'table'; }
export function isDiagramElement(el: NormalizedElement): el is NormalizedDiagramElement { return el.type === 'diagram'; }
export function isVideoElement(el: NormalizedElement): el is NormalizedVideoElement { return el.type === 'video'; }
export function isConnectionElement(el: NormalizedElement): el is NormalizedConnectionElement { return el.type === 'connection'; }

export type RelationshipGraph = Record<string, unknown>;
export type MediaFiles = Record<string, Uint8Array>;

export interface NormalizedResult {
  format: NormalizedFormat;
  slides: NormalizedSlide[];
  /**
   * How the slide order was decided: 'presentation' means p:sldIdLst was read,
   * 'filename' means it could not be and creation order was assumed.
   */
  slideOrderSource?: 'presentation' | 'filename';
  slideDimensions?: { width: number; height: number };
  mediaFiles: MediaFiles;
  relationships: RelationshipGraph;
  slideLayoutRelationships?: Record<string, string>;
  theme?: {
    colors: Record<string, string>;
    rawTheme: any;
  };
}
