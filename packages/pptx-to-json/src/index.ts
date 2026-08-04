/**
 * @cliftonc/pptx-to-json — main exports
 *
 * This package contains all the PowerPoint parsing logic and processors.
 */

// Main processors
export { PowerPointClipboardProcessor } from './processors/PowerPointClipboardProcessor.js';
export { PPTXParser } from './processors/PPTXParser.js';
export type { PartRelationship } from './processors/PPTXParser.js';

// Individual parsers
export { BaseParser } from './parsers/BaseParser.js';
export { PowerPointParser } from './parsers/PowerPointParser.js';
export { PowerPointNormalizer } from './parsers/PowerPointNormalizer.js';
export { TextParser } from './parsers/TextParser.js';
export { ShapeParser } from './parsers/ShapeParser.js';
export { ImageParser } from './parsers/ImageParser.js';
export { TableParser } from './parsers/TableParser.js';
export { VideoParser } from './parsers/VideoParser.js';
export { DiagramParser } from './parsers/DiagramParser.js';
export { ConnectorParser } from './parsers/ConnectorParser.js';
export { SlideMetaParser } from './parsers/SlideMetaParser.js';
export type { SlideTitle } from './parsers/SlideMetaParser.js';

// Component types — the shape of what a parse returns
export type {
  ComponentType,
  PowerPointComponent,
  PowerPointComponentBase,
  TextComponent,
  ShapeComponent,
  ImageComponent,
  TableComponent,
  TableRow,
  TableCell,
  DiagramComponent,
  VideoComponent,
  ConnectionComponent,
  UnknownComponent,
  ComponentStyle,
  TextRun,
  FillInfo,
  BorderInfo,
  GeometryInfo,
  EffectsInfo,
  PlaceholderMap,
  PlaceholderPosition,
  SmartArtDataPoint,
  SmartArtConnection,
  SmartArtShape,
  SmartArtLayout,
  XMLNode
} from './types/index.js';

// Normalized (intermediate) structure — slides, their order, titles and notes
export type {
  NormalizedFormat,
  NormalizedResult,
  NormalizedSlide,
  NormalizedElement
} from './types/normalized.js';

// Parse results
export type {
  ParsedResult,
  ParsedSlide,
  ParsedLayout,
  ParsedMaster,
  SlideMetadata
} from './parsers/PowerPointParser.js';
