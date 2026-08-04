/**
 * Slide structure: presentation order, titles and speaker notes.
 *
 * These go through the real path — zip bytes → PPTXParser.buffer2json → PowerPointParser
 * — rather than hand-built json, because two of the three answers depend on how the XML
 * parser is configured. In particular `r:id` must stay distinguishable from `id`; a test
 * against a pre-parsed tree would keep passing if `removeNSPrefix` were ever switched on.
 */

import { describe, it, expect } from 'vitest'
import { PPTXParser } from '../../src/processors/PPTXParser.ts'
import { PowerPointParser } from '../../src/parsers/PowerPointParser.ts'
import { SlideMetaParser } from '../../src/parsers/SlideMetaParser.ts'
import { buildDeck, fakePng, type DeckSpec } from '../helpers/deck-builder.ts'

async function parseDeck(spec: DeckSpec) {
  const buffer = await buildDeck(spec)
  const pptxParser = new PPTXParser()
  const json = await pptxParser.buffer2json(buffer)
  const parser = new PowerPointParser()
  return { result: await parser.parseJson(json), json, pptxParser }
}

describe('slide order', () => {
  it('follows p:sldIdLst, not the slide filenames', async () => {
    // A deck reordered in PowerPoint: part names keep their creation order
    const { result } = await parseDeck({
      slides: [
        { file: 1, title: { text: 'Created first' } },
        { file: 2, title: { text: 'Created second' } },
        { file: 3, title: { text: 'Created third' } }
      ],
      presentationOrder: [3, 1, 2]
    })

    expect(result.slideOrderSource).toBe('presentation')
    expect(result.slides.map(s => s.title)).toEqual([
      'Created third',
      'Created first',
      'Created second'
    ])

    // slideNumber is the position in the deck; the filename number is kept separately
    expect(result.slides.map(s => s.slideNumber)).toEqual([1, 2, 3])
    expect(result.slides.map(s => s.metadata.fileSlideNumber)).toEqual([3, 1, 2])
    expect(result.slides.map(s => s.metadata.slideFile)).toEqual([
      'ppt/slides/slide3.xml',
      'ppt/slides/slide1.xml',
      'ppt/slides/slide2.xml'
    ])
  })

  it('sorts by slide number, not lexically, past ten slides', async () => {
    const files = Array.from({ length: 12 }, (_, i) => i + 1)
    const { result } = await parseDeck({
      slides: files.map(file => ({ file, title: { text: `Slide ${file}` } })),
      presentationOrder: files
    })

    expect(result.slides.map(s => s.title)).toEqual(files.map(f => `Slide ${f}`))
  })

  it('falls back to filename order, and says so, when there is no presentation part', async () => {
    const { result } = await parseDeck({
      slides: [
        { file: 2, title: { text: 'Second' } },
        { file: 1, title: { text: 'First' } }
      ],
      omitPresentation: true
    })

    expect(result.slideOrderSource).toBe('filename')
    expect(result.slides.map(s => s.title)).toEqual(['First', 'Second'])
  })

  it('resolves relationship targets against the part that declares them', () => {
    expect(
      PPTXParser.resolveRelationshipTarget(
        'ppt/slides/_rels/slide1.xml.rels',
        '../notesSlides/notesSlide1.xml'
      )
    ).toBe('ppt/notesSlides/notesSlide1.xml')

    expect(
      PPTXParser.resolveRelationshipTarget('ppt/_rels/presentation.xml.rels', 'slides/slide4.xml')
    ).toBe('ppt/slides/slide4.xml')

    expect(
      PPTXParser.resolveRelationshipTarget('ppt/_rels/presentation.xml.rels', '/ppt/slides/slide4.xml')
    ).toBe('ppt/slides/slide4.xml')
  })
})

describe('speaker notes', () => {
  it('attaches notes through the rels graph, not by matching numbers', async () => {
    // The trap: notesSlide parts are numbered in their own sequence, so slide 2 owns
    // notesSlide1 as soon as slide 1 has no notes at all
    const { result } = await parseDeck({
      slides: [
        { file: 1, title: { text: 'No notes here' } },
        { file: 2, title: { text: 'Second' }, notes: { file: 1, text: 'Notes belonging to slide 2' } },
        { file: 3, title: { text: 'Third' }, notes: { file: 2, text: 'Notes belonging to slide 3' } }
      ]
    })

    expect(result.slides.map(s => s.notes)).toEqual([
      undefined,
      'Notes belonging to slide 2',
      'Notes belonging to slide 3'
    ])
    expect(result.slides.map(s => s.metadata.notesFile)).toEqual([
      null,
      'ppt/notesSlides/notesSlide1.xml',
      'ppt/notesSlides/notesSlide2.xml'
    ])
  })

  it('reports no notes for a part holding only the auto-generated slide number', async () => {
    // Most notes parts in real decks look exactly like this
    const { result } = await parseDeck({
      slides: [{ file: 1, title: { text: 'Slide' }, notes: { file: 1, includeSlideNumberField: true } }]
    })

    expect(result.slides[0].notes).toBeUndefined()
  })

  it('keeps notes with the slide they belong to when the deck is reordered', async () => {
    const { result } = await parseDeck({
      slides: [
        { file: 1, title: { text: 'A' }, notes: { file: 1, text: 'Notes for A' } },
        { file: 2, title: { text: 'B' }, notes: { file: 2, text: 'Notes for B' } }
      ],
      presentationOrder: [2, 1]
    })

    expect(result.slides.map(s => ({ title: s.title, notes: s.notes }))).toEqual([
      { title: 'B', notes: 'Notes for B' },
      { title: 'A', notes: 'Notes for A' }
    ])
  })
})

describe('slide titles', () => {
  it('reads a title placeholder, including ctrTitle', async () => {
    const { result } = await parseDeck({
      slides: [
        { file: 1, title: { text: 'Centre title', type: 'ctrTitle' } },
        { file: 2, title: { text: 'Ordinary title', type: 'title' } }
      ]
    })

    expect(result.slides.map(s => s.title)).toEqual(['Centre title', 'Ordinary title'])
    expect(result.slides.map(s => s.metadata.titleSource)).toEqual(['placeholder', 'placeholder'])
  })

  it('falls back to the layout when the placeholder declares no type', async () => {
    const { result } = await parseDeck({
      slides: [{ file: 1, untypedPlaceholder: { text: 'Inherited title', idx: 0 }, body: 'Body text' }],
      layoutPlaceholders: { 0: 'title', 1: 'body' }
    })

    expect(result.slides[0].title).toBe('Inherited title')
    expect(result.slides[0].metadata.titleSource).toBe('layout')
  })

  it('leaves the title undefined when the layout says the placeholder is a body', async () => {
    const { result } = await parseDeck({
      slides: [{ file: 1, untypedPlaceholder: { text: 'Just body copy', idx: 1 } }],
      layoutPlaceholders: { 0: 'title', 1: 'body' }
    })

    expect(result.slides[0].title).toBeUndefined()
    expect(result.slides[0].metadata.titleSource).toBeNull()
  })

  it('has no title when the slide has no title placeholder', async () => {
    const { result } = await parseDeck({ slides: [{ file: 1, body: 'Only body text' }] })

    expect(result.slides[0].title).toBeUndefined()
  })

  it('keeps a title that is entirely numeric', async () => {
    // The XML parser types a run of "2026" as a number; a string-only read drops it
    const { result } = await parseDeck({ slides: [{ file: 1, title: { text: '2026' } }] })

    expect(result.slides[0].title).toBe('2026')
  })
})

describe('media stays with its own slide', () => {
  it('resolves images through the slide rels part, not slide number arithmetic', async () => {
    const { result } = await parseDeck({
      slides: [
        { file: 1, title: { text: 'First created' }, image: 'imageA.png' },
        { file: 2, title: { text: 'Second created' }, image: 'imageB.png' }
      ],
      // Reversed: slide2.xml is presented first, so `slideNumber` and the rels part
      // it would imply disagree
      presentationOrder: [2, 1],
      media: { 'imageA.png': fakePng('MARKER-A'), 'imageB.png': fakePng('MARKER-B') }
    })

    const imageOf = (index: number) => {
      const image = result.slides[index].components.find(c => c.type === 'image') as any
      const base64 = String(image?.src ?? '').split(',')[1] ?? ''
      return Buffer.from(base64, 'base64').toString('binary')
    }

    expect(result.slides[0].title).toBe('Second created')
    expect(imageOf(0)).toContain('MARKER-B')
    expect(imageOf(1)).toContain('MARKER-A')
  })
})

describe('tables', () => {
  it('keeps the text of every cell', async () => {
    // Cell text used to come back empty from every table in every deck: the extractor
    // called `require` in an ESM package and the ReferenceError was swallowed
    const { result } = await parseDeck({
      slides: [
        {
          file: 1,
          title: { text: 'With a table' },
          table: [
            ['Region', 'Spend'],
            ['EMEA', '23,700'],
            ['APAC', '18,200']
          ]
        }
      ]
    })

    const table = result.slides[0].components.find(c => c.type === 'table') as any
    expect(table).toBeDefined()
    expect(table.rows.map((row: any) => row.cells.map((cell: any) => cell.content))).toEqual([
      ['Region', 'Spend'],
      ['EMEA', '23,700'],
      ['APAC', '18,200']
    ])
  })
})

describe('SlideMetaParser', () => {
  it('reads nothing from an absent tree rather than throwing', () => {
    expect(SlideMetaParser.extractTitle(null)).toBeNull()
    expect(SlideMetaParser.extractNotes(undefined)).toBeNull()
  })
})
