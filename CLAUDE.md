# PowerPoint Paste Parser - Implementation Summary

## Development Workflow

You can use the log-bin command on either an extracted clipboard file (these are stored in `packages/pptx-to-json/test/test-harness/fixtures`), or any pptx file.

```
pnpm log-bin <filename>        # Parse binary file from fixtures/
pnpm log-bin bullets.pptx       # Example: parse bullets.bin
pnpm log-bin bullets.pptx --slide 1       # Example: limit output to slide 1
pnpm log-bin bullets.bin --debug  # Example: with debug output
```

To test changes you are making you simply re-run the log-bin to see the output and it should be modified.

## Key Files

**`/packages/pptx-to-json/src/index.ts`** - Library exports for the main server component and TypeScript interfaces.

**`/packages/pptx-to-json/src/parsers`** - Individual parsers for sections of powerpoint (text, shape, image, video, table)

**`/packages/pptx-to-json/src/processors`** - Core processor for PPTX (or clipboard).  Clipboard content is actually the same structure, just with a different namespace - solved via the normalizer - and without a slide master / layout.

### Raw vs namespace-stripped JSON — the one rule that bites

`PowerPointNormalizer.normalize()` strips every namespace prefix (`p:spPr` → `spPr`) and
almost all parsing works on that stripped tree. **The relationship graph does not.**
`p:sldId` carries both an `id` and an `r:id`, and stripping merges them, so slide order would
resolve correctly on most decks and silently wrongly on others. Anything reading `r:id` —
`PPTXParser.getSlideOrder`, `getSlideNotesRelationships`, `getPartRelationships` — takes the
RAW json, which `normalize()` passes through alongside the stripped one. For the same reason,
never set `removeNSPrefix: true` on the `XMLParser`.

Two consequences worth remembering:
- **Slide filenames are creation order.** `slideNumber` is the position in `p:sldIdLst`;
  `metadata.fileSlideNumber` is the number in the filename, and that is what the rels graph is
  keyed on. Anything resolving a slide's media takes `slideRelsFile`, never `slideIndex + 1`.
- **Notes are matched through the rels graph**, never `slideN` ↔ `notesSlideM`.

**`/packages/ppt-paste-parser/src/ClipboardParser.tsx`** - Main React component that handles clipboard paste events, detects PowerPoint cloud service metadata, extracts Microsoft API URLs, and calls the worker server to get parsed components.

**`/apps/worker/worker.js`** - Cloudflare Worker with Hono framework that serves the client app and provides API endpoints. Bypasses CORS restrictions, calls Microsoft's GetClipboardBytes API, and uses the pptx-to-json package for parsing.

**`/apps/worker/client/src/App.tsx`** - Client application that displays parsed PowerPoint components with TLDraw integration, slideshow mode, and structured UI with visual styling.

**`/packages/ppt-paste-parser/src/index.ts`** - Library exports for the ClipboardParser component and TypeScript interfaces.

## Project Overview

PowerPoint Component Parser that extracts structured data from PowerPoint files and clipboard paste operations. For pasting from Clipboard, the system intercepts PowerPoint paste operations, calls Microsoft's internal APIs, and parses the Office Open XML format to provide detailed component analysis.

## System Architecture

### Server-Side Library (`packages/pptx-to-json/`)
- Full parsing of both pptx and clipboard files
- This is where key logic and complexity lives

### Client-Side Library (`packages/ppt-paste-parser/`)
- Minimal implementation that detects PowerPoint data, calls worker server, and returns structured components
- React component with TypeScript interfaces
- Delegates parsing to the server

### Cloudflare Worker (`apps/worker/`)
- Hono-based worker running on Cloudflare Workers platform
- Serves static client application via assets binding
- Provides API endpoints for PowerPoint parsing
- Uses pptx-to-json package for ZIP/XML processing
- Supports R2 storage for uploaded PPTX files
- Returns structured JSON with parsed components

### Client Application (`apps/worker/client/`)
- React application with Vite build system
- TLDraw integration for visual presentation canvas
- Slideshow mode with keyboard navigation
- Displays PowerPoint component data with visual styling
- Built to `/apps/worker/dist` for worker deployment

## What Gets Parsed

The system extracts:

1. Component Types: text, image, shape, table, unknown
2. Positions: x,y coordinates from PowerPoint
3. Dimensions: Width and height in PowerPoint units
4. Content: Text content from text components
5. Metadata: Additional component information
6. IDs: Unique component identifiers

## File Structure

```
ppt-paste/
├── packages/
│   ├── ppt-paste-parser/               # Client library (TypeScript/React)
│   │   ├── src/ClipboardParser.tsx     # Main component (clean, minimal)
│   │   ├── src/index.ts                # Exports
│   │   └── dist/                       # Built library
│   └── pptx-to-json/               # Server parsing logic
│       └── src/                        # ZIP/XML parsing, PowerPoint parsers
│           └── parsers/                # invididual element parsing
│           └── processors/             # Processing of pptx or clipboard
├── apps/worker/                        # Cloudflare Worker app
│   ├── client/                         # Client application
│   │   ├── src/                        # React components, TLDraw integration
│   │   ├── package.json                # Client dependencies
│   │   └── vite.config.ts              # Builds to ../dist
│   ├── dist/                           # Built client (deployment target)
│   ├── worker.js                       # Cloudflare Worker with Hono
│   ├── package.json                    # Worker dependencies and scripts
│   └── wrangler.toml                   # Cloudflare configuration
└── CLAUDE.md                           # This summary
```

## How to Run

### Development
```bash
# Start both worker and client in development
pnpm dev

# This runs:
# - Worker: http://localhost:3001 (API + serves built client)
# - Client dev server: http://localhost:5173 (with hot reload)
```

### Production Build & Deploy
```bash
# Build everything and deploy to Cloudflare
pnpm deploy
```

### Usage
- Copy content from PowerPoint (shapes, text, images)
- Paste in the application (dev: http://localhost:5173, prod: your-worker.workers.dev)
- View component parsing with TLDraw visualization
- Use slideshow mode for presentation view

## Technical Implementation

### Microsoft API Integration
- Reverse-engineered PowerPoint's cloud service integration
- Calls `GetClipboardBytes.ashx` API endpoint
- Handles authentication and CORS via Cloudflare Worker

### Office Open XML Processing
- ZIP extraction from binary clipboard data or pptx file using jszip library
- PowerPoint drawing XML format parsing with specialized component parsers
- Component positions, sizes, and content extraction with modular parser architecture
- Proper handling of ppt format (`a:txSp` or `p:txSp` structure) vs full file format

### Parser Architecture

**Core Parsers** (`/packages/pptx-to-json/src/parsers/`):
- **BaseParser.js**: Common utilities (EMU conversion, color parsing, transform extraction)
- **PowerPointParser.js**: Main coordinator, handles clipboard vs file format detection
- **TextParser.js**: Text components with comprehensive font/style parsing
- **ShapeParser.js**: Geometric shapes with fill, border, and effect parsing
- **ImageParser.js**: Images with data URL extraction and effects
- **TableParser.js**: Tables with cell structure and formatting
- **SlideMetaParser.ts**: Slide titles and speaker notes — the text that describes a slide rather than drawing on it

### Architecture
- **Modular Parser System**: Specialized parsers for different component types (TextParser, ShapeParser, ImageParser, TableParser)
- **Clipboard Format Support**: Proper detection and handling of PowerPoint clipboard structure (`a:txSp` → `a:txBody`)
- **Enhanced Component Classification**: Reliable text vs shape detection with fallback logic
- **Type-safe interfaces**: TypeScript throughout the stack
- **Server-side processing**: Client delegates parsing to Cloudflare Worker for better performance
- **Edge deployment**: Runs on Cloudflare's global edge network for low latency

## Development Commands

```bash
# Development
pnpm dev                        # Start worker + client dev servers

# Building
pnpm build                      # Build library + client app
pnpm deploy                     # Build + deploy to Cloudflare Workers

# Type checking
pnpm type-check
```

## Dependencies

**Client Library:**
- React 19+
- TypeScript

**Cloudflare Worker:**
- Hono (web framework)
- pptx-to-json (parsing logic)
- Cloudflare Workers Runtime
- R2 storage (for uploaded files)

**Client App:**
- Vite (build tool)
- React 19+
- TLDraw (canvas rendering)
- TipTap (rich text editing)
- ppt-paste-parser (local)

---
