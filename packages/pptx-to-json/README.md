# @cliftonc/pptx-to-json

A TypeScript library for parsing PowerPoint files (PPTX) and PowerPoint clipboard data into structured JSON. Extracts components like text, shapes, images, and tables with their positioning, styling, and content information.

## 🌐 Live Demo

Try it out at **[ppt-paste.clifton-cunningham.workers.dev](https://ppt-paste.clifton-cunningham.workers.dev/)**

## Features

- **PPTX File Parsing**: Parse complete PowerPoint files
- **Clipboard Data Parsing**: Handle PowerPoint clipboard paste operations
- **Component Extraction**: Extract text, shapes, images, tables, and more
- **Position & Styling**: Get accurate positioning, dimensions, and styling information
- **Slide structure**: presentation order, titles and speaker notes, resolved through the
  relationship graph rather than guessed from filenames (see [Slide structure](#slide-structure))
- **TypeScript Support**: Full TypeScript definitions included
- **Modular Architecture**: Extensible parser system for different component types

## Installation

```bash
npm install @cliftonc/pptx-to-json
```

## Usage

### Parsing PPTX Files

```typescript
import { PPTXParser, PowerPointParser } from '@cliftonc/pptx-to-json';
import fs from 'node:fs';

const buffer = fs.readFileSync('presentation.pptx');

// 1. unzip + parse every part to JSON
const json = await new PPTXParser().buffer2json(buffer);

// 2. turn that into slides and components
const result = await new PowerPointParser().parseJson(json);

for (const slide of result.slides) {
  console.log(slide.slideNumber, slide.title);   // presentation order, title placeholder
  console.log(slide.notes);                      // speaker notes, if any
  console.log(slide.components.length);
}
```

### Parsing Clipboard Data

```typescript
import { PowerPointClipboardProcessor } from '@cliftonc/pptx-to-json';

const processor = new PowerPointClipboardProcessor();
const result = await processor.parseClipboardBuffer(clipboardBuffer);

console.log(result.slides);
```

## Slide structure

Three answers are resolved through the package relationship graph, because the obvious
shortcut is wrong in ways that are hard to notice afterwards:

| | Read from | Why not the obvious way |
|---|---|---|
| `slide.slideNumber` | `ppt/presentation.xml` → `p:sldIdLst` → `r:id` | Slide **filenames are creation order**. PowerPoint keeps a slide's part name when a deck is reordered, so `slide3.xml` can be the seventh slide shown. `result.slideOrderSource` reports whether this list could be read; `slide.metadata.fileSlideNumber` keeps the filename number. |
| `slide.notes` | `ppt/slides/_rels/slideN.xml.rels` → `notesSlide` | A notes part exists only for slides that HAVE notes, so `notesSlideM` stops lining up with `slideN` as soon as one slide lacks them. Notes attributed to the wrong slide are worse than no notes at all. Only the body placeholder is read: nearly every real notes part contains just the auto-generated slide-number field, which would otherwise be returned as "notes". |
| `slide.title` | `p:ph type="title"\|"ctrTitle"`, else the layout's placeholder for that `idx` | A slide's placeholder often omits `@type` and is named only by the layout it inherits from. `slide.metadata.titleSource` says which of the two answered. |

> ⚠️ **Do not set `removeNSPrefix: true` on the `XMLParser`.** It merges `id` and `r:id`
> into one key, and which one wins depends on attribute order in the source — slide-order
> resolution would appear to work on most decks and silently mis-resolve on others. The
> parser is configured with `ignoreAttributes: false` and `attributeNamePrefix: "$"`, and
> does **not** strip prefixes. Keep it that way.

# PowerPoint Parse Output Format Specification

This document provides a comprehensive specification for the structured JSON produced by the PowerPoint parsing pipeline (`pptx` files and clipboard paste data) in this repository. It consolidates details from:
- Normalized intermediate types (`src/types/normalized.ts`)
- Public component types (`src/types/index.ts`)
- Parser implementations (`TextParser`, `ShapeParser`, `ImageParser`, `TableParser`, `DiagramParser`, `VideoParser`, `ConnectorParser`)
- Aggregation logic in `PowerPointParser`
- Existing test fixtures in `test/test-harness/expected/*.json`

The goal is to give consumers a stable contract for what a "parse result" contains and how to interpret each field.

---
## 1. High-Level Result Shapes

Two conceptual layers exist:
1. Normalized Intermediate (`NormalizedResult`) – internal, z-order–preserving structure used during parsing.
2. Parsed Result (`ParsedResult`) – external structure of slides, masters, layouts, and components returned by `PowerPointParser.parseJson` (and downstream helpers).

Depending on which public API you call (e.g. legacy wrappers like `parsePptx` / `parseClipboard` if reintroduced), you will typically receive a `ParsedResult`-like object or a simpler wrapper exposing the `slides` array.

### 1.1 ParsedResult (Primary External Shape)
```
interface ParsedResult {
  slides: ParsedSlide[];            // In presentation order
  masters: Record<string, ParsedMaster>; // Keyed by master XML path (ppt/slideMasters/slideMasterX.xml)
  layouts: Record<string, ParsedLayout>; // Keyed by layout XML path (ppt/slideLayouts/slideLayoutY.xml)
  totalComponents: number;          // Total count across all slides (after expansions)
  format: 'pptx' | 'clipboard';     // Source format detected by normalizer
  slideOrderSource?: 'presentation' | 'filename'; // Whether p:sldIdLst decided the order
  slideDimensions?: { width: number; height: number }; // In pixels (EMU converted)
}
```

### 1.2 ParsedSlide
```
interface ParsedSlide {
  slideIndex: number;        // Zero-based position in presentation order
  slideNumber: number;       // One-based position in presentation order
  title?: string;            // Title placeholder text, when the slide has one
  notes?: string;            // Speaker notes, when the slide has any
  layoutId?: string;         // Derived from layout file name (slideLayoutX)
  background?: PowerPointComponent; // Background image/shape (if non-white & accepted)
  components: PowerPointComponent[]; // All foreground components, including background IF pushed into flow earlier
  metadata: SlideMetadata;   // Provenance + counts (see below)
}

interface SlideMetadata {
  name: string;                    // The slide title, or 'Slide N' when it has none
  componentCount: number;          // components.length
  format: 'pptx' | 'clipboard';
  slideFile: string | null;        // e.g. 'ppt/slides/slide1.xml'
  layoutFile: string | null;       // e.g. 'ppt/slideLayouts/slideLayout1.xml'
  masterFile: string | null;       // e.g. 'ppt/slideMasters/slideMaster1.xml'
  layoutElementCount: number;      // Count of inherited layout elements (pre-filter)
  masterElementCount: number;      // Count of inherited master elements (pre-filter)
  notesFile: string | null;        // e.g. 'ppt/notesSlides/notesSlide2.xml'
  fileSlideNumber: number | null;  // The number in the filename — creation order
  titleSource: 'placeholder' | 'layout' | null; // How the title was identified
}
```

### 1.3 ParsedMaster
```
interface ParsedMaster {
  id: string;                     // e.g. 'slideMaster1'
  name: string;                   // 'Master 1'
  background?: PowerPointComponent; // Non-white fill or image background
  components: PowerPointComponent[]; // Master-level decorative or thematic components
  sourceFile: string;             // Full master path
  placeholders?: PlaceholderMap;  // Placeholder geometry by idx or type:*
  textStyles?: {                  // Extracted txStyles for font inheritance
    titleStyle?: XMLNode;
    bodyStyle?: XMLNode;
  };
}
```

### 1.4 ParsedLayout
```
interface ParsedLayout {
  id: string;                      // e.g. 'slideLayout3'
  name: string;                    // 'Layout 3'
  masterId?: string;               // Associated master id
  background?: PowerPointComponent;
  components: PowerPointComponent[];
  sourceFile: string;              // layout file path
  placeholders?: PlaceholderMap;   // Layout-level placeholders override masters
}
```

---
## 2. Component Model

All returned components are discriminated by `type`. Base fields:
```
interface PowerPointComponentBase {
  id: string;              // Unique within result; generated: <Kind or Name> <OriginalName or Index>
  type: ComponentType;     // 'text' | 'shape' | 'image' | 'table' | 'video' | 'connection' | 'diagram' | 'any'
  content?: string;        // Human-readable summary or plain extracted text
  x: number;               // Pixel position (EMU converted)
  y: number;
  width: number;           // Pixels
  height: number;          // Pixels
  rotation?: number;       // Degrees; may be 0; negative supported
  style?: ComponentStyle;  // See Style section
  metadata?: Record<string, any>; // Parser + provenance details
  slideIndex: number;      // Owning slide zero-based index (masters/layouts use 0 when synthesized)
  zIndex: number;          // Higher means on top; backgrounds use large negative or normalized values
}
```

### 2.1 TextComponent
```
interface TextComponent extends PowerPointComponentBase {
  type: 'text';
  richText?: RichTextDoc;              // TipTap-like structure (paragraphs & bulletLists)
  backgroundShape?: {                  // Only present for text-with-visible background geometry
    type: 'rectangle' | 'ellipse' | 'roundRect' | 'custom';
    fill?: FillInfo;                   // { type, color, opacity }
    border?: BorderInfo;               // { type, color, width, style, cap?, compound? }
    geometry?: GeometryInfo;           // Normalized geometry description
  };
}
```
Key text-specific metadata fields:
- `metadata.isTextBox`: boolean (derived from `cNvSpPr.$txBox`)
- `metadata.paragraphCount`: number
- `metadata.hasMultipleRuns`: boolean (true if any paragraph has multiple runs)
- `metadata.namespace`: `'a' | 'p'` indicates clipboard vs pptx source segment

Spacing heuristics: spaces may be synthesized between adjacent alphanumeric runs when PowerPoint omits them across run boundaries.

### 2.2 ShapeComponent
```
interface ShapeComponent extends PowerPointComponentBase {
  type: 'shape';
  shapeType?: string;        // Human-friendly descriptor (e.g. 'rectangle', 'up arrow')
  geometry?: GeometryInfo;   // { type, preset, isCustom, paths? }
}
```
Style fields usually populated:
- `style.fillColor`, `style.fillOpacity`
- `style.borderColor`, `style.borderWidth`, `style.borderStyle`
- Effects (if any) merged into style: possible keys include `effects`, `boxShadow`, etc.

Metadata booleans:
- `hasEffects`, `hasFill`, `hasBorder`

### 2.3 ImageComponent
```
interface ImageComponent extends PowerPointComponentBase {
  type: 'image';
  src?: string;      // Data URL or externally hosted URL (if R2 / remote)
  alt?: string;      // Description from cNvPr.$descr
}
```
Metadata additions:
- `relationshipId`: underlying rId
- `imageUrl`: dup of `src` for clarity
- `imageType`: normalized extension (png, jpeg, gif, etc.)
- `imageSize`: byte length (if known)
- `hasEffects`: boolean

Notes:
- WMF/EMF images are skipped (treated as decorative backgrounds).
- Background images use negative z-index only during intermediate; accepted backgrounds appear in slide `background` or as first component in flow depending on processing stage.

### 2.4 TableComponent
```
interface TableComponent extends PowerPointComponentBase {
  type: 'table';
  rows?: { cells: { content: string; style?: ComponentStyle; colSpan?: number; rowSpan?: number }[] }[];
  columns?: number; // Count of columns
}
```
Metadata:
- `tableData`: 2D string matrix (raw textual extraction)
- `rows`, `cols`: numeric dimensions
- `hasHeader`: currently `true` if first row is treated as header in `richText` structure
- `richText`: TipTap table representation mirroring client expectation
- `format`: namespace `'a' | 'p'`

### 2.5 DiagramComponent
```
interface DiagramComponent extends PowerPointComponentBase {
  type: 'diagram';
  diagramType?: string;        // e.g. 'smartart' or 'unknown'
  title?: string;              // Derived fallback name
  smartArtData?: {             // Only for SmartArt when fully extracted
    dataPoints: SmartArtDataPoint[];    // Logical model nodes
    connections: SmartArtConnection[];  // Logical edges
    shapes: SmartArtShape[];            // Visual shape placements
    layout: SmartArtLayout;             // Layout metadata
  };
  extractedComponents?: PowerPointComponent[]; // Flattened shapes/text derived from SmartArt
}
```
If `options.returnExtractedComponents` is set during parse, the original `DiagramComponent` wrapper may be replaced by its `extractedComponents` array (each item a standard `PowerPointComponent` type – often `shape` or `text`).

### 2.6 VideoComponent
```
interface VideoComponent extends PowerPointComponentBase {
  type: 'video';
  url?: string;                 // Resolved target from slide relationships
  thumbnailSrc?: string;        // Data URL for poster image (if available)
  title?: string;               // From cNvPr.$name
  embedType?: 'youtube' | 'vimeo' | 'generic';
}
```
Resolution logic distinguishes between PPTX mode (relationship XML) and clipboard mode (simpler relationship map). If no match is found, `url` may be `undefined`.

### 2.7 ConnectionComponent (Connectors / Lines)
```
interface ConnectionComponent extends PowerPointComponentBase {
  type: 'connection';
  startShapeId?: string;         // Target shape ID (parsed from connection metadata)
  endShapeId?: string;
  connectorType?: string;        // Derived line archetype
  startPoint?: { x: number; y: number }; // Computed anchor point in pixels
  endPoint?: { x: number; y: number };
  lineStyle?: {                  // Extracted from <ln>
    width?: number;              // Pixels
    color?: string;              // Hex color
    dashStyle?: string;          // e.g. 'dash'
    startArrow?: string;         // Arrowhead descriptor
    endArrow?: string;
  };
}
```
Post-processing (`fixConnectionPoints`) refines `startPoint` / `endPoint` after all shapes are known, using indices stored in `metadata.startConnectionIndex` / `metadata.endConnectionIndex` and bounding boxes of associated shapes.

### 2.8 UnknownComponent
```
interface UnknownComponent extends PowerPointComponentBase {
  type: 'any';
}
```
Rare fallback when a parser cannot classify an element but still returns structural coordinates.

---
## 3. Styles & Visual Data

```
interface ComponentStyle {
  fontFamily?: string;
  fontSize?: number;         // Points for text; numeric px-like scale after extraction
  fontWeight?: string;       // 'bold', 'normal', etc.
  fontStyle?: string;        // 'italic', 'normal'
  textAlign?: string;        // 'left' | 'center' | 'right' | 'justify'
  color?: string;            // Foreground text color (hex)
  borderWidth?: number;      // Pixels
  borderStyle?: string;      // 'solid' | 'none' | dash patterns
  borderColor?: string;
  fillColor?: string;        // Background fill (shape/table)
  fillOpacity?: number;      // 0–1 normalized
  rotation?: number;         // Degrees
  // Additional dynamic keys include effects, shadows, filters
  [key: string]: any;
}
```

Fill, border, geometry helpers (from shape parsing):
```
interface FillInfo { type: 'solid' | 'gradient' | 'pattern' | 'none'; color: string; opacity: number; }
interface BorderInfo { type: 'solid' | 'none'; color: string; width: number; style: string; cap?: string; compound?: string; }
interface GeometryInfo { type: string; preset: string | null; isCustom: boolean; paths?: any[]; }
```

Effects (`EffectsInfo` – merged into style) may include:
- `effects: string[]`
- `boxShadow?: string`
- Additional effect-specific keys

Image effects augmented fields:
- `opacity`, `filter`, `borderRadius`, `shadow`, `effectsList`

---
## 4. Coordinate & Unit Handling

- PowerPoint internally uses EMUs (English Metric Units).
- All exported `x`, `y`, `width`, `height` are converted to **pixels** using centralized constants (`emuToPixels` / `validatePixelRange`).
- Rotation is converted from 1/60000 degree units to plain degrees.
- Slide dimensions (`slideDimensions`) are also pixel-converted.
- Guard rails: any dimension > 50,000 triggers a warning (likely unconverted EMU scenario).

---
## 5. Z-Ordering & Background Logic

- Backgrounds extracted from masters/layouts/slides are given strongly negative `zIndex` during normalization:
  - Master backgrounds: start at -2000 and decrement.
  - Layout backgrounds: start at -1000 and decrement.
  - Slide-specific backgrounds: promoted with `zIndex` of -500 when parsed and may be stored in both `slide.background` and `components` (depending on stage) if they pass the non-white criteria.
- White / near-white solid fills are suppressed (ignored) for master/layout backgrounds to avoid noise.
- Image backgrounds are always preserved.

---
## 6. Placeholder & Master Style Inheritance

Text and shape components with zero transform size/position will look up placeholder geometry if:
1. They reference a placeholder (`nvSpPr > nvPr > ph`), and
2. A matching placeholder (`idx` or `type`) exists in merged placeholder maps (master first, then layout override).

Master text styles (`titleStyle`, `bodyStyle`) affect font size inheritance when a text element lacks explicit `rPr` size declarations.

Priority order for font size resolution:
1. First explicit non-empty run with size
2. Placeholder-derived master `titleStyle` (if placeholder type is 'title')
3. Master `bodyStyle`
4. Default fallback (Arial 18, or other parser default)

---
## 7. Rich Text Representation

All rich text surfaces (text boxes and table cells) are expressed in a simplified TipTap-like JSON:
```
{ type: 'doc', content: Array<Paragraph | BulletList | TableNode> }
```
Supported nodes:
- `paragraph` – sequential inline `text` nodes
- `bulletList` → `listItem` → nested `paragraph`
- `table` (for tables only) → `tableRow` → `tableCell` / `tableHeader` → `paragraph`

Text marks (`marks`) include:
- `{ type: 'bold' }`
- `{ type: 'italic' }`
- `{ type: 'textStyle', attrs: { fontSize: '22pt', color: '#FFFFFF', ... } }`

Spacing heuristics maintain readability while avoiding double insertion of whitespace.

---
## 8. Diagram (SmartArt) Expansion

When SmartArt is detected:
- `diagramType` set to `smartart`
- Raw relationship IDs are scanned to collect companion diagram/data/drawing files.
- `smartArtData` includes:
  - `dataPoints`: logical node set with hierarchy & content
  - `connections`: edges linking data points
  - `shapes`: positioned shapes representing each data node
  - `layout`: summarizing layoutType, category, colorScheme, quickStyle
- If `returnExtractedComponents` flag is provided, the parser returns the flattened `extractedComponents` instead of the diagram wrapper, allowing uniform downstream handling.

---
## 9. Connections (Lines/Arrows)

Connectors rely on:
- `startConnection` / `endConnection` from normalized XML (IDs + connection site indices)
- A second pass to compute pixel anchor points (center, mid-left, mid-right, etc.) using bounding boxes of referenced shapes
- Style extraction from `<ln>` (width, color, dash style, arrowheads)

Metadata fields assist post-processing:
- `startConnectionIndex`, `endConnectionIndex`
- Boolean flags `hasStartConnection`, `hasEndConnection`

---
## 10. Media & Relationship Resolution

Images & videos use relationship graphs:
- PPTX mode: relationships are file-based (`ppt/slides/_rels/slideN.xml.rels`), and the
  slide's own rels part is passed down explicitly. Deriving it from the slide's position
  would read another slide's relationships once a deck has been reordered.
- Clipboard mode: relationships stored in a simplified map keyed by slide index.
- Background images restructure `blipFill` to surface `embed` id for uniform handling.
- Video thumbnails resolve through nested `blipFill` -> `embed` in `thumbnailRelationshipId` if present.

Skipped media:
- `wmf` / `emf` images (often vector placeholders or noise).

---
## 11. Background Acceptance Rules

A candidate background is added only if:
- It is an image, OR
- It has a non-white/near-white fill (solid, gradient, pattern).

White-ish backgrounds are suppressed to reduce redundancy, especially on layout/master inheritance, letting slide content show default canvas background.

---
## 12. IDs & Names

`id` strategy uses `BaseParser.generateComponentId(kind, componentIndex, originalName)` ensuring uniqueness. Connectors append semi-colon–delimited segments for later shape ID extraction. Connection component `id` format example:
```
<GeneratedConnectorName>;<UnderlyingShapeId>;<namespace><slideIndex>
```
Shape IDs referenced by connectors are extracted from the second `;` segment (`shapeId`).

---
## 13. Error Handling & Null Returns

Each specialized parser returns `null` to skip components when:
- Missing structural nodes (`spPr`, `blipFill`, `graphicData`, etc.)
- Zero width/height (for shapes, diagrams)
- Empty text content (text parser)
- Filtered image types (wmf/emf)

The aggregator silently ignores `null` (optionally logging in debug mode).

---
## 14. Examples

### 14.1 Text Component (Clipboard)
```
{
  "id": "TextBox 4",
  "type": "text",
  "content": "TEXT",
  "x": 109,
  "y": 41,
  "width": 288,
  "height": 139,
  "style": { "fontSize": 80, "color": "#4EA72E", ... },
  "richText": { "type": "doc", "content": [ { "type": "paragraph", ... } ] },
  "backgroundShape": { "type": "rectangle", "border": { ... }, "geometry": { ... } },
  "metadata": { "namespace": "a", "paragraphCount": 1 }
}
```

### 14.2 Shape Component
```
{
  "id": "Arrow: Up 12",
  "type": "shape",
  "content": "up arrow shape",
  "x": 341,
  "y": 471,
  "width": 124,
  "height": 176,
  "style": { "fillColor": "#4472C4", "borderStyle": "none", ... },
  "shapeType": "up arrow",
  "geometry": { "type": "up arrow", "preset": "upArrow", "isCustom": false },
  "metadata": { "hasFill": true, "hasBorder": false }
}
```

### 14.3 Image Component (Background)
```
{
  "id": "Google Shape;308;p58",
  "type": "image",
  "content": "Google Shape;308;p58",
  "x": 25,
  "y": 53,
  "width": 569,
  "height": 473,
  "src": "data:image/png;base64,...",
  "metadata": { "relationshipId": "rId1", "imageType": "png", "imageSize": 220591 }
}
```

### 14.4 Connection Component
```
{
  "id": "Connector 5;42;p0",
  "type": "connection",
  "content": "Connector 5 Connection",
  "x": 100, "y": 120, "width": 200, "height": 0,
  "startShapeId": "1001",
  "endShapeId": "1002",
  "startPoint": { "x": 300, "y": 140 },
  "endPoint": { "x": 100, "y": 140 },
  "lineStyle": { "width": 2, "color": "#000000", "dashStyle": "solid" },
  "metadata": { "startConnectionIndex": 3, "endConnectionIndex": 1 }
}
```

(Example generalized; actual IDs may differ.)

---
## 15. Versioning & Stability Notes

- The discriminated union of `PowerPointComponent` is the primary stable contract.
- New component types (e.g. future media or interactive elements) will extend the union with a distinct `type` value and additive fields.
- Optional metadata fields are additive; consumers should feature-detect.
- Normalized internal structures are subject to refinement but will continue to feed the same external component shapes.

---
## 16. Migration & Compatibility Guidelines

When integrating:
1. **Rely on `type`** for branching logic.
2. **Use presence checks** (`if ('richText' in component)`) before accessing optional fields.
3. **Do not assume** backgrounds appear only in `slide.background` – they may also appear as the first `components[]` entry for legacy consumers.
4. **Expect arrays** like `richText.content` and `rows` to be empty rather than `null` when no data exists.
5. **Ignore unknown keys** in `style` – future effect enrichments will appear there.

---
## 17. Known Limitations / TODO

- SmartArt extraction: Coverage may vary; complex layouts might produce partial `smartArtData`.
- Table styling: Cell-level font/color details are currently minimal (structure prioritized over rich formatting).
- Video thumbnails: Not always available depending on relationship resolution.
- Connectors: Complex routing (elbows, curves) simplified to bounding box endpoints; no path geometry yet.
- Diagram expansion may return large component sets; consider filtering `extractedComponents` downstream.


## Development

This library is part of a larger PowerPoint parsing ecosystem. The complete system includes:

- **[@cliftonc/pptx-to-json](https://www.npmjs.com/package/@cliftonc/pptx-to-json)**: This parsing library
- **Web Application**: Live demo and visual interface
- **Cloudflare Worker**: API endpoints for parsing operations

### Local Development

```bash
# Build the library
npm run build

# Run tests
npm test

# Type checking
npm run type-check

# Test with sample files
npm run log-bin sample.pptx
```

## Technical Details

### Architecture

The library uses a modular parser architecture:

- **BaseParser**: Common utilities for EMU conversion, colors, transforms
- **PowerPointParser**: Main coordinator handling format detection
- **TextParser**: Text components with font/style parsing
- **ShapeParser**: Geometric shapes with fill, border, effects
- **ImageParser**: Images with data URL extraction
- **TableParser**: Tables with cell structure and formatting

### Format Support

- **PPTX Files**: Complete Office Open XML PowerPoint files
- **Clipboard Data**: PowerPoint clipboard format with proper namespace handling
- **Component Detection**: Reliable classification with fallback logic

## License

MIT

## Contributing

Issues and pull requests are welcome! Please visit the [GitHub repository](https://github.com/cliftonc/pptx-to-json) for more information.

## Related Projects

- 🌐 **[Live Demo](https://ppt-paste.clifton-cunningham.workers.dev/)**: Try the parser in your browser
- 📚 **Documentation**: Full implementation details in the main repository

---

Made with ❤️ by [Clifton Cunningham](https://github.com/cliftonc)
