/**
 * Slide-level metadata: the things that describe a slide rather than draw on it.
 *
 * Titles and speaker notes are both plain text pulled out of a placeholder, and both
 * have to skip the same auto-generated furniture (slide-number fields, footers), so
 * they share one text reader here rather than growing two that drift apart.
 *
 * ⚠️ Every method takes a NAMESPACE-STRIPPED node (`sp`, `txBody`, `p`, `r`, `t`), which
 * is what `PowerPointNormalizer.stripNamespaces` produces. The rels graph — where `r:id`
 * and `id` must stay distinguishable — is resolved in `PPTXParser` against the raw tree.
 */

import type { XMLNode } from '../types/index.js';

/** Placeholder types that are page furniture, never body content */
const FURNITURE_PLACEHOLDERS = new Set(['sldNum', 'dt', 'ftr', 'hdr', 'sldImg']);

/** Placeholder types that make a shape the slide's title */
const TITLE_PLACEHOLDERS = new Set(['title', 'ctrTitle']);

export interface SlideTitle {
  text: string;
  /**
   * 'placeholder' — the slide's own `ph` said `title`/`ctrTitle`.
   * 'layout'      — the slide's `ph` carried only an `idx`, and the layout named it.
   */
  source: 'placeholder' | 'layout';
}

export class SlideMetaParser {
  /**
   * The slide's title, or null when it has no title placeholder (27% of real slides).
   *
   * @param spTree - namespace-stripped `cSld/spTree` of the slide
   * @param layoutPlaceholderTypes - idx → type from the slide's layout (see
   *   `PPTXParser.getPlaceholderTypesByIdx`). A slide placeholder often carries only an
   *   `idx` and inherits its type, so without this a titled slide reads as untitled.
   */
  static extractTitle(
    spTree: XMLNode | null | undefined,
    layoutPlaceholderTypes: Record<string, string> = {}
  ): SlideTitle | null {
    if (!spTree) return null;

    for (const sp of SlideMetaParser.toArray(spTree['sp'])) {
      const ph = sp?.['nvSpPr']?.['nvPr']?.['ph'];
      if (!ph) continue;

      const declaredType = ph.$type ?? ph.type;
      let source: SlideTitle['source'];

      if (declaredType !== undefined && declaredType !== null) {
        if (!TITLE_PLACEHOLDERS.has(String(declaredType))) continue;
        source = 'placeholder';
      } else {
        // No @type means the type is inherited; @idx defaults to 0 when absent
        const idx = String(ph.$idx ?? ph.idx ?? 0);
        const inherited = layoutPlaceholderTypes[idx];
        if (!inherited || !TITLE_PLACEHOLDERS.has(inherited)) continue;
        source = 'layout';
      }

      // A title laid out over several paragraphs is still one line of text
      const text = SlideMetaParser.readText(sp['txBody'], ' ');
      if (text) return { text, source };
    }

    return null;
  }

  /**
   * Speaker notes for a slide, or null when there are none.
   *
   * Most decks carry a notes part for nearly every slide and real notes on almost none:
   * the part is usually an empty body placeholder plus an auto-generated slide-number
   * field. Reading everything in the part would therefore return the slide number as
   * "notes" for most slides, so only the body placeholder is read and fields are skipped.
   *
   * @param notesSlide - namespace-stripped notes part (the whole file node, or its `notes` element)
   */
  static extractNotes(notesSlide: XMLNode | null | undefined): string | null {
    if (!notesSlide) return null;

    const notes = notesSlide['notes'] ?? notesSlide;
    const spTree = notes?.['cSld']?.['spTree'];
    if (!spTree) return null;

    const parts: string[] = [];

    for (const sp of SlideMetaParser.toArray(spTree['sp'])) {
      const ph = sp?.['nvSpPr']?.['nvPr']?.['ph'];
      const type = ph ? String(ph.$type ?? ph.type ?? 'body') : null;

      // Shapes with no placeholder at all are hand-added notes content; keep them
      if (type !== null && FURNITURE_PLACEHOLDERS.has(type)) continue;

      const text = SlideMetaParser.readText(sp['txBody'], '\n');
      if (text) parts.push(text);
    }

    const combined = parts.join('\n').trim();
    return combined.length > 0 ? combined : null;
  }

  /**
   * Read the text of a `txBody`, joining paragraphs with `paragraphSeparator`.
   *
   * Two details that are easy to lose:
   * - `a:fld` runs are SKIPPED. They hold auto-generated values — slide numbers, dates —
   *   which are not authored content, and a slide number is exactly what an unfiltered
   *   read returns for an otherwise empty notes part.
   * - Run text is coerced with `String()`. `parseTagValue` turns a run whose whole text
   *   is numeric into a NUMBER, so a title of "2026" is not a string and a `typeof
   *   text === 'string'` guard drops it without a trace.
   */
  private static readText(txBody: XMLNode | null | undefined, paragraphSeparator: string): string {
    if (!txBody) return '';

    const paragraphs: string[] = [];

    for (const paragraph of SlideMetaParser.toArray(txBody['p'])) {
      if (!paragraph || typeof paragraph !== 'object') continue;

      let text = '';
      for (const run of SlideMetaParser.toArray(paragraph['r'])) {
        const value = run?.['t'];
        if (value === undefined || value === null) continue;
        text += typeof value === 'object' ? String(value._text ?? '') : String(value);
      }

      // A line break inside a paragraph is still a space in a flattened title
      if (paragraph['br'] && text) text += ' ';

      const trimmed = text.trim();
      if (trimmed) paragraphs.push(trimmed);
    }

    return paragraphs.join(paragraphSeparator).replace(/[ \t]+/g, ' ').trim();
  }

  private static toArray(value: any): any[] {
    if (value === undefined || value === null) return [];
    return Array.isArray(value) ? value : [value];
  }
}
