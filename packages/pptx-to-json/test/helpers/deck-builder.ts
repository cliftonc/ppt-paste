/**
 * Builds synthetic .pptx files in memory.
 *
 * Real decks are the evidence, but they cannot be committed, so every test here works
 * from a deck assembled part by part — which also lets a test express the situation that
 * matters and is hard to find in the wild: a deck whose slides have been REORDERED, so
 * that presentation order and file numbering disagree.
 */

import JSZip from 'jszip'

const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface SlideSpec {
  /** The slide's file number — `ppt/slides/slide<n>.xml`. This is CREATION order. */
  file: number
  /** A title placeholder declaring its own type */
  title?: { text: string; type?: 'title' | 'ctrTitle' }
  /** A placeholder carrying only an idx, whose type has to come from the layout */
  untypedPlaceholder?: { text: string; idx?: number }
  /** Body text in an ordinary (non-placeholder) text box */
  body?: string
  /** An image, related as rId20, pointing at `ppt/media/<image>` */
  image?: string
  /** A table, as rows of cell text */
  table?: string[][]
  /** Notes part number — `ppt/notesSlides/notesSlide<n>.xml` */
  notes?: { file: number; text?: string; includeSlideNumberField?: boolean }
}

export interface DeckSpec {
  slides: SlideSpec[]
  /** Slide file numbers in PRESENTATION order. Defaults to the order `slides` is given in. */
  presentationOrder?: number[]
  /** Media parts, keyed by filename under `ppt/media/` */
  media?: Record<string, Uint8Array>
  /** idx → placeholder type declared by the layout */
  layoutPlaceholders?: Record<number, string>
  /** Omit ppt/presentation.xml entirely, to exercise the filename-order fallback */
  omitPresentation?: boolean
}

const xmlDecl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

function textBody(text: string): string {
  return `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB" dirty="0"/><a:t>${text}</a:t></a:r></a:p></p:txBody>`
}

function shapeProps(x: number, y: number): string {
  return `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="5486400" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr>`
}

function slideXml(spec: SlideSpec): string {
  const shapes: string[] = []
  let id = 2

  if (spec.title) {
    shapes.push(
      `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Title"/><p:cNvSpPr/>` +
        `<p:nvPr><p:ph type="${spec.title.type ?? 'title'}"/></p:nvPr></p:nvSpPr>` +
        `${shapeProps(0, 0)}${textBody(spec.title.text)}</p:sp>`
    )
  }

  if (spec.untypedPlaceholder) {
    // No @type: the type is inherited from the layout, by idx
    shapes.push(
      `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Placeholder"/><p:cNvSpPr/>` +
        `<p:nvPr><p:ph idx="${spec.untypedPlaceholder.idx ?? 0}"/></p:nvPr></p:nvSpPr>` +
        `${shapeProps(0, 0)}${textBody(spec.untypedPlaceholder.text)}</p:sp>`
    )
  }

  if (spec.body) {
    shapes.push(
      `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `${shapeProps(0, 2000000)}${textBody(spec.body)}</p:sp>`
    )
  }

  const pic = spec.image
    ? `<p:pic><p:nvPicPr><p:cNvPr id="${id++}" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId20"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
    : ''

  let table = ''
  if (spec.table) {
    const grid = (spec.table[0] ?? []).map(() => '<a:gridCol w="2743200"/>').join('')
    const rows = spec.table
      .map(
        row =>
          `<a:tr h="370840">${row
            .map(
              cell =>
                `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-GB"/><a:t>${cell}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
            )
            .join('')}</a:tr>`
      )
      .join('')

    table =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id++}" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
      `<p:xfrm><a:off x="0" y="3000000"/><a:ext cx="5486400" cy="2000000"/></p:xfrm>` +
      `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
      `<a:tbl><a:tblPr firstRow="1"/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl>` +
      `</a:graphicData></a:graphic></p:graphicFrame>`
  }

  return `${xmlDecl}<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `${shapes.join('')}${pic}${table}</p:spTree></p:cSld></p:sld>`
}

function slideRelsXml(spec: SlideSpec): string {
  const rels = [
    `<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`
  ]
  if (spec.notes) {
    rels.push(
      `<Relationship Id="rId2" Type="${REL}/notesSlide" Target="../notesSlides/notesSlide${spec.notes.file}.xml"/>`
    )
  }
  if (spec.image) {
    rels.push(`<Relationship Id="rId20" Type="${REL}/image" Target="../media/${spec.image}"/>`)
  }
  return `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`
}

function notesXml(notes: NonNullable<SlideSpec['notes']>): string {
  // The slide-number field is what nearly every real notes part contains and nothing else
  const slideNumberShape = notes.includeSlideNumberField
    ? `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/>` +
      `<p:txBody><a:bodyPr/><a:p><a:fld id="{GUID}" type="slidenum"><a:rPr lang="en-GB"/><a:t>7</a:t></a:fld></a:p></p:txBody></p:sp>`
    : ''

  const body =
    `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder"/><p:cNvSpPr/>` +
    `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
    (notes.text
      ? textBody(notes.text)
      : '<p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="en-GB"/></a:p></p:txBody>') +
    '</p:sp>'

  return `${xmlDecl}<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `${body}${slideNumberShape}</p:spTree></p:cSld></p:notes>`
}

function layoutXml(placeholders: Record<number, string>): string {
  const shapes = Object.entries(placeholders).map(
    ([idx, type], i) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${i + 2}" name="ph${idx}"/><p:cNvSpPr/>` +
      `<p:nvPr><p:ph type="${type}" idx="${idx}"/></p:nvPr></p:nvSpPr>` +
      `${shapeProps(0, 0)}<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`
  )

  return `${xmlDecl}<p:sldLayout xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" type="obj"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `${shapes.join('')}</p:spTree></p:cSld></p:sldLayout>`
}

function masterXml(): string {
  return `${xmlDecl}<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `</p:spTree></p:cSld><p:txStyles><p:titleStyle/><p:bodyStyle/></p:txStyles></p:sldMaster>`
}

/** Build a .pptx buffer from the spec */
export async function buildDeck(spec: DeckSpec): Promise<Uint8Array> {
  const zip = new JSZip()
  const order = spec.presentationOrder ?? spec.slides.map(s => s.file)

  if (!spec.omitPresentation) {
    // rId1 is the master, so slide relationship ids deliberately do NOT line up with
    // slide numbers — the same offset a real deck has
    const slideIds = order
      .map((fileNumber, i) => `<p:sldId id="${256 + i}" r:id="rId${slideRelId(fileNumber, order)}"/>`)
      .join('')

    zip.file(
      'ppt/presentation.xml',
      `${xmlDecl}<p:presentation xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">` +
        `<p:sldIdLst>${slideIds}</p:sldIdLst>` +
        `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`
    )

    const presRels = order
      .map(
        fileNumber =>
          `<Relationship Id="rId${slideRelId(fileNumber, order)}" Type="${REL}/slide" Target="slides/slide${fileNumber}.xml"/>`
      )
      .join('')

    zip.file(
      'ppt/_rels/presentation.xml.rels',
      `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
        `${presRels}</Relationships>`
    )
  }

  for (const slide of spec.slides) {
    zip.file(`ppt/slides/slide${slide.file}.xml`, slideXml(slide))
    zip.file(`ppt/slides/_rels/slide${slide.file}.xml.rels`, slideRelsXml(slide))
    if (slide.notes) {
      zip.file(`ppt/notesSlides/notesSlide${slide.notes.file}.xml`, notesXml(slide.notes))
    }
  }

  zip.file('ppt/slideLayouts/slideLayout1.xml', layoutXml(spec.layoutPlaceholders ?? {}))
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `${xmlDecl}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`
  )
  zip.file('ppt/slideMasters/slideMaster1.xml', masterXml())

  for (const [name, bytes] of Object.entries(spec.media ?? {})) {
    zip.file(`ppt/media/${name}`, bytes)
  }

  return zip.generateAsync({ type: 'uint8array' })
}

/** Relationship id for a slide: offset past rId1 (the master) by its position in the list */
function slideRelId(fileNumber: number, order: number[]): number {
  return order.indexOf(fileNumber) + 2
}

/** A distinguishable PNG-ish blob: a real signature so type detection works, a unique tail */
export function fakePng(marker: string): Uint8Array {
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
  const tail = Array.from(new TextEncoder().encode(marker))
  return new Uint8Array([...signature, ...tail])
}
