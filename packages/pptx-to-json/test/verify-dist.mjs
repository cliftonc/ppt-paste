/**
 * Smoke check on the built package: can a consumer actually import it?
 *
 * `dist/` is gitignored and the build is incremental, so a stale
 * tsconfig.tsbuildinfo makes tsc emit only the files that changed and leaves the
 * rest missing. Nothing catches that: type-check passes, the tests run from src,
 * `npm pack` happily lists whatever is there, and the first sign of trouble is
 * ERR_MODULE_NOT_FOUND in a consumer's app. `prepublishOnly` runs this.
 */

const REQUIRED_EXPORTS = [
  'PPTXParser',
  'PowerPointParser',
  'PowerPointNormalizer',
  'PowerPointClipboardProcessor',
  'BaseParser',
  'TextParser',
  'ShapeParser',
  'ImageParser',
  'TableParser',
  'VideoParser',
  'DiagramParser',
  'ConnectorParser',
  'SlideMetaParser'
]

const built = await import('../dist/index.js')

const missing = REQUIRED_EXPORTS.filter(name => typeof built[name] !== 'function')
if (missing.length > 0) {
  console.error(`dist/index.js is missing exports: ${missing.join(', ')}`)
  console.error('The build is incomplete — remove dist/ and tsconfig.tsbuildinfo and rebuild.')
  process.exit(1)
}

// Exercise the entry point rather than only its shape: a module can import fine
// and still be half a package
const parser = new built.PPTXParser()
if (typeof parser.getSlideOrder !== 'function') {
  console.error('dist is stale: PPTXParser has no getSlideOrder')
  process.exit(1)
}

console.log(`dist ok — ${REQUIRED_EXPORTS.length} exports resolve and load`)
