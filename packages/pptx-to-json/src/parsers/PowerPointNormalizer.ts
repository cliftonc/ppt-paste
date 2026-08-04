/**
 * PowerPoint Structure Normalizer
 * 
 * Converts both PPTX files and clipboard data to a unified internal structure,
 * eliminating the need for if-statements throughout the parsing logic.
 */

import { PPTXParser } from '../processors/PPTXParser.js';
import { DEFAULT_SLIDE_WIDTH_PX, DEFAULT_SLIDE_HEIGHT_PX, pixelsToEmu } from '../utils/constants.js';
import { BaseParser } from './BaseParser.js';
import { SlideMetaParser } from './SlideMetaParser.js';

import type { NormalizedElement, NormalizedSlide, NormalizedResult } from '../types/normalized.js';
/* Inlined NormalizedElement/Slide/Result interfaces removed.
   Using shared types from ../types/normalized.ts */

export class PowerPointNormalizer {
  
  /**
   * Normalize PowerPoint JSON data to a unified structure
   */
  normalize(json: any): NormalizedResult {
    const formatType = this.detectFormat(json);

    // Strip all namespaces from the entire JSON structure at the very start
    const strippedJson = this.stripNamespaces(json);

    switch (formatType) {
      case 'pptx':
        // The RAW json is passed alongside: the rels graph (slide order, notes) is read
        // from it, because stripping namespaces makes `r:id` indistinguishable from `id`
        return this.normalizePPTX(strippedJson, formatType, json);
      case 'clipboard':
        return this.normalizeClipboard(strippedJson, formatType);
      default:
        throw new Error(`Unknown PowerPoint format: ${formatType}`);
    }
  }
  
  /**
   * Detect the format type based on file structure
   */
  detectFormat(json: any): string {
    const files = Object.keys(json);
    
    // Check for PPTX structure
    if (files.some(f => f.startsWith('ppt/slides/'))) {
      return 'pptx';
    }
    
    // Check for clipboard structure  
    if (files.some(f => f.startsWith('clipboard/drawings/'))) {
      return 'clipboard';
    }
    
    throw new Error('Unable to detect PowerPoint format - no recognized file structure found');
  }
  
  /**
   * Normalize PPTX file structure
   */
  normalizePPTX(json: any, formatType: string, rawJson?: any): NormalizedResult {
    const slides: NormalizedSlide[] = [];
    const files = Object.keys(json);

    // Create PPTXParser instance for layout extraction
    const pptxParser = new PPTXParser();

    // Get slide-to-layout relationships
    const slideLayoutRelationships = pptxParser.getSlideLayoutRelationships(json);

    // Extract slide dimensions from presentation.xml
    const slideDimensions = pptxParser.getSlideDimensions(json);

    // Extract theme data from theme file
    const themeData = this.extractThemeData(json, 'pptx');

    // The rels graph is read from the RAW tree — see normalize() above. Callers that
    // reach normalizePPTX directly with only a stripped tree still work, they just fall
    // back to filename order and get no notes.
    const relsSource = rawJson ?? json;
    const slideOrder = pptxParser.getSlideOrder(relsSource);
    const notesRelationships = pptxParser.getSlideNotesRelationships(relsSource);
    const slideOrderSource: 'presentation' | 'filename' = slideOrder.length > 0 ? 'presentation' : 'filename';
    const orderByFile = new Map(slideOrder.map((file, index) => [file, index]));

    // idx → placeholder type per layout, so a slide placeholder carrying only an idx
    // can still be recognised as the title. Parsed once per layout, not once per slide.
    const layoutPlaceholderTypes = new Map<string, Record<string, string>>();
    const placeholderTypesFor = (layoutFile: string | undefined): Record<string, string> => {
      if (!layoutFile) return {};
      if (!layoutPlaceholderTypes.has(layoutFile)) {
        layoutPlaceholderTypes.set(layoutFile, pptxParser.getPlaceholderTypesByIdx(json, layoutFile));
      }
      return layoutPlaceholderTypes.get(layoutFile) as Record<string, string>;
    };

    // Find slide files (no sorting needed - we'll extract slide numbers)
    const slideFiles = files.filter(f =>
      f.startsWith('ppt/slides/slide') && f.endsWith('.xml')
    );

    for (const slideFile of slideFiles) {
      const slideData = json[slideFile];
      if (!slideData || !slideData['sld']) continue;

      // Extract slide number from the FILENAME (e.g., 'ppt/slides/slide5.xml' -> 5).
      // This is creation order, not presentation order — the two are reconciled below.
      const slideNumberMatch = slideFile.match(/slide(\d+)\.xml/);
      const fileSlideNumber = slideNumberMatch ? parseInt(slideNumberMatch[1], 10) : 1;
      const slideNumber = fileSlideNumber;

      // Extract slide content
      const slide = slideData['sld'];
      const cSld = this.ensureObject(slide['cSld']) || slide;
      const spTree = this.ensureObject(cSld['spTree']);
      
      if (!spTree) continue;
      
      // Extract slide elements
      const slideElements = this.extractOrderedElements(spTree);
      
      // Check for slide-specific background
      const slideBackground = cSld['bg'];
      if (slideBackground) {
        const bgPr = slideBackground['bgPr'] || slideBackground;
        const blipFill = bgPr && (bgPr['blipFill'] || bgPr['a:blipFill']);
        
        // console.log(`Slide ${slideNumber} background check:`, {
        //   hasBackground: !!slideBackground,
        //   hasBgPr: !!bgPr,
        //   hasBlipFill: !!blipFill,
        //   bgPrKeys: bgPr ? Object.keys(bgPr) : [],
        //   slideFile
        // });
        
        // Determine if this is an image background or shape background
        if (blipFill) {
          // Create an image background element
          const backgroundElement: NormalizedElement = {
            type: 'image',
            namespace: 'p',
            element: 'pic',
            data: {
              nvPicPr: {
                cNvPr: {
                  $id: -500,
                  $name: 'Slide Background Image'
                }
              },
              blipFill,
              spPr: {
                xfrm: {
                  off: { $x: 0, $y: 0 },
                  ext: { $cx: pixelsToEmu(slideDimensions.width), $cy: pixelsToEmu(slideDimensions.height) } // Use actual slide dimensions
                }
              },
              isBackground: true
            },
            zIndex: -500,
            isBackgroundElement: true,
            blipFill // Add blipFill as direct property for easier access
          };
          slideElements.push(backgroundElement);
        } else {
          // Create a shape background element (solid color, gradient, etc.)
          const backgroundElement: NormalizedElement = {
            type: 'shape',
            namespace: 'p',
            element: 'sp',
            data: {
              nvSpPr: {
                cNvPr: {
                  $id: -500,
                  $name: 'Slide Background'
                }
              },
              spPr: bgPr,
              isBackground: true
            },
            zIndex: -500,
            isBackgroundElement: true
          };
          slideElements.push(backgroundElement);
        }
      }
      
      // Get layout elements for this slide
      const layoutFile = slideLayoutRelationships[slideFile];
      let layoutElements: NormalizedElement[] = [];
      let masterElements: NormalizedElement[] = [];
      let masterFile: string | null = null;
      
      if (layoutFile) {
        // Get layout elements
        const rawLayoutElements = pptxParser.getSlideLayoutElements(json, layoutFile, slideDimensions);
        
        // Convert layout elements to the same structure as slide elements
        layoutElements = rawLayoutElements.map(layoutEl => {
          const convertedElement = this.convertLayoutElementToSlideFormat(layoutEl, spTree);
          return convertedElement;
        }).filter(el => el !== null) as NormalizedElement[];
        
        // Get master elements for this layout
        const layoutMasterRelationships = pptxParser.getLayoutMasterRelationships(json);
        masterFile = layoutMasterRelationships[layoutFile];
        
        if (masterFile) {
          const rawMasterElements = pptxParser.getSlideMasterElements(json, masterFile, slideDimensions);
          
          // Convert master elements to the same structure as slide elements
          masterElements = rawMasterElements.map(masterEl => {
            const convertedElement = this.convertMasterElementToSlideFormat(masterEl, spTree);
            return convertedElement;
          }).filter(el => el !== null) as NormalizedElement[];
        }
      }
      
      // Merge elements in proper z-order hierarchy:
      // 1. Master elements (deepest background, z-index -2000 to -1999)
      // 2. Layout elements (middle background, z-index -1000 to -999) 
      // 3. Slide elements (foreground, z-index 0+)
      const allElements = [...masterElements, ...layoutElements, ...slideElements];
      
      const notesFile = notesRelationships[slideFile];
      const notes = notesFile ? SlideMetaParser.extractNotes(json[notesFile]) : null;
      const title = SlideMetaParser.extractTitle(spTree, placeholderTypesFor(layoutFile));

      const normalizedSlide: NormalizedSlide = {
        slideFile,
        slideNumber, // Presentation position, assigned after sorting below
        fileSlideNumber, // The number in the filename — creation order
        slideRelsFile: `ppt/slides/_rels/slide${fileSlideNumber}.xml.rels`,
        title: title?.text,
        titleSource: title?.source,
        notes: notes ?? undefined,
        notesFile: notesFile ?? null,
        format: 'pptx',
        shapes: this.extractPPTXShapes(spTree),
        images: this.extractPPTXImages(spTree),
        text: this.extractPPTXText(spTree),
        videos: this.extractPPTXVideos(spTree),
        elements: allElements, // Combined master, layout and slide elements
        layoutFile, // Keep track of which layout this slide uses
        masterFile: masterFile || null, // Keep track of which master this slide uses
        layoutElementCount: layoutElements.length,
        masterElementCount: masterElements.length,
        rawSpTree: spTree // Keep for relationship lookups
      };
      
      slides.push(normalizedSlide);
    }

    // Put the slides in the order the deck is presented in.
    //
    // Until now they were in whatever order the zip listed them, and their number came
    // from the filename. Both are creation order: PowerPoint keeps a slide's part name
    // when the deck is reordered, so slide3.xml can be the seventh slide shown. The
    // authoritative order is p:sldIdLst (see PPTXParser.getSlideOrder); filename number
    // is the fallback when there is no readable presentation part, and is kept as
    // `fileSlideNumber` either way because the rels graph is keyed on it.
    slides.sort((a, b) => {
      const aOrder = orderByFile.get(a.slideFile);
      const bOrder = orderByFile.get(b.slideFile);
      // Slides absent from sldIdLst are not presented; keep them last, in filename order
      if (aOrder === undefined || bOrder === undefined) {
        if (aOrder === undefined && bOrder === undefined) {
          return (a.fileSlideNumber ?? 0) - (b.fileSlideNumber ?? 0);
        }
        return aOrder === undefined ? 1 : -1;
      }
      return aOrder - bOrder;
    });

    slides.forEach((slide, index) => {
      slide.slideNumber = index + 1;
    });

    return {
      format: 'pptx',
      slides,
      slideOrderSource,
      slideDimensions,
      mediaFiles: this.extractMediaFiles(json),
      relationships: this.extractRelationships(json),
      slideLayoutRelationships, // Include relationships for reference
      theme: themeData // Include theme data
    };
  }
  
  /**
   * Normalize clipboard structure
   */
  normalizeClipboard(json: any, formatType: string): NormalizedResult {
    const slides: NormalizedSlide[] = [];
    const files = Object.keys(json);
    
    // Extract theme data from clipboard theme file
    const themeData = this.extractThemeData(json, 'clipboard');
    
    // Find drawing files
    const drawingFiles = files.filter(f => 
      f.includes('clipboard') && f.includes('drawing') && f.endsWith('.xml')
    );
    
    for (const drawingFile of drawingFiles) {
      const drawingData = json[drawingFile];
      if (!drawingData || !drawingData['graphic']) continue;
      
      // Navigate clipboard structure
      const graphic = drawingData['graphic'];
      const graphicData = graphic['graphicData'];
      if (!graphicData) continue;
      
      const lockedCanvas = graphicData['lockedCanvas'];
      if (!lockedCanvas) continue;
      
      const normalizedSlide: NormalizedSlide = {
        slideFile: drawingFile,
        format: 'clipboard',
        shapes: this.extractClipboardShapes(lockedCanvas),
        images: this.extractClipboardImages(lockedCanvas), 
        text: this.extractClipboardText(lockedCanvas),
        videos: this.extractClipboardVideos(lockedCanvas),
        elements: this.extractOrderedClipboardElements(lockedCanvas), // Add ordered elements
        rawCanvas: lockedCanvas // Keep for relationship lookups
      };
      
      slides.push(normalizedSlide);
    }
    
    // Calculate content-based dimensions for clipboard format
    const contentBounds = this.calculateContentBounds(slides);
    
    return {
      format: 'clipboard',
      slides,
      slideDimensions: contentBounds,
      mediaFiles: this.extractMediaFiles(json),
      relationships: this.extractRelationships(json),
      theme: themeData // Include theme data
    };
  }
  
  /**
   * Extract shapes from PPTX spTree
   */
  extractPPTXShapes(spTree: any): any[] {
    const shapes = [];
    const spArray = this.ensureArray(spTree['sp']);
    
    for (const sp of spArray) {
      // Skip text boxes (they're handled in extractText)
      if (sp['nvSpPr'] && sp['nvSpPr']['cNvSpPr'] && sp['nvSpPr']['cNvSpPr']['$txBox']) continue;
      
      // Skip shapes with actual text content (they're handled in extractPPTXText)
      if (sp['txBody'] && this.hasTextContent(sp['txBody'])) continue;
      
      // Check if this shape uses background fill (design element)
      const useBgFill = sp['$useBgFill'] === '1' || sp['useBgFill'] === '1' || sp['$useBgFill'] === 1 || sp['useBgFill'] === 1;
      
      
      // Skip shapes that use background fill as they are typically design elements
      if (useBgFill) continue;
      
      shapes.push({
        type: 'shape',
        namespace: 'p',
        element: 'sp',
        data: sp,
        spPr: sp['spPr'],
        nvSpPr: sp['nvSpPr'],
        style: sp['style'], // Add style data for fill/border parsing
        textBody: sp['txBody'] // Direct access for empty text containers
      });
    }
    
    return shapes;
  }
  
  /**
   * Extract shapes from clipboard lockedCanvas
   */
  extractClipboardShapes(lockedCanvas: any): any[] {
    const shapes = [];
    const spData = lockedCanvas['sp'];
    
    if (spData) {
      // Clipboard usually has single shape object, not array
      const spArray = this.ensureArray(spData);
      
      for (const sp of spArray) {
        // Skip text boxes (they're handled in extractText)
        if (sp['nvSpPr'] && sp['nvSpPr']['cNvSpPr'] && sp['nvSpPr']['cNvSpPr']['$txBox']) continue;
        
        // Skip shapes with actual text content (they're handled in extractClipboardText)
        if (sp['txSp'] && sp['txSp']['txBody'] && this.hasTextContent(sp['txSp']['txBody'])) continue;
        
        shapes.push({
          type: 'shape',
          namespace: 'a',
          element: 'sp', 
          data: sp,
          spPr: sp['spPr'],
          nvSpPr: sp['nvSpPr'],
          style: sp['style'], // Add style data for fill/border parsing
          textBody: sp['txSp'] ? sp['txSp']['txBody'] : null // Extra layer!
        });
      }
    }
    
    return shapes;
  }
  
  /**
   * Extract images from PPTX spTree
   */
  extractPPTXImages(spTree: any): any[] {
    const images = [];
    const picArray = this.ensureArray(spTree['pic']);
    
    for (const pic of picArray) {
      // Skip videos - they have nvPr.videoFile property
      const nvPicPr = pic['nvPicPr'];
      if (nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile']) {
        continue; // This is a video, not an image
      }
      
      images.push({
        type: 'image',
        namespace: 'p',
        element: 'pic',
        data: pic,
        nvPicPr: pic['nvPicPr'],
        blipFill: pic['blipFill'],
        spPr: pic['spPr']
      });
    }
    
    return images;
  }
  
  /**
   * Extract images from clipboard lockedCanvas
   */
  extractClipboardImages(lockedCanvas: any): any[] {
    const images = [];
    const picData = lockedCanvas['pic'];
    
    if (picData) {
      const picArray = this.ensureArray(picData);
      
      for (const pic of picArray) {
        // Skip videos - they have nvPr.videoFile property
        const nvPicPr = pic['nvPicPr'];
        if (nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile']) {
          continue; // This is a video, not an image
        }
        
        images.push({
          type: 'image',
          namespace: 'a',
          element: 'pic',
          data: pic,
          nvPicPr: pic['nvPicPr'], 
          blipFill: pic['blipFill'],
          spPr: pic['spPr']
        });
      }
    }
    
    return images;
  }
  
  /**
   * Extract videos from PPTX spTree
   */
  extractPPTXVideos(spTree: any): any[] {
    const videos = [];
    const picArray = this.ensureArray(spTree['pic']);
    
    for (const pic of picArray) {
      // Check if this is a video by looking for nvPr.videoFile
      const nvPicPr = pic['nvPicPr'];
      if (nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile']) {
        const videoFile = nvPicPr['nvPr']['videoFile'];
        const relationshipId = videoFile['$link'] || videoFile['link'];
        
        videos.push({
          type: 'video',
          namespace: 'p',
          element: 'pic',
          data: pic,
          nvPicPr: pic['nvPicPr'],
          blipFill: pic['blipFill'],
          spPr: pic['spPr'],
          relationshipId: relationshipId
        });
      }
    }
    
    return videos;
  }
  
  /**
   * Extract videos from clipboard lockedCanvas
   */
  extractClipboardVideos(lockedCanvas: any): any[] {
    const videos = [];
    const picData = lockedCanvas['pic'];
    
    if (picData) {
      const picArray = this.ensureArray(picData);
      
      for (const pic of picArray) {
        // Check if this is a video by looking for nvPr.videoFile
        const nvPicPr = pic['nvPicPr'];
        if (nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile']) {
          const videoFile = nvPicPr['nvPr']['videoFile'];
          const relationshipId = videoFile['$link'] || videoFile['link'];
          
          videos.push({
            type: 'video',
            namespace: 'a',
            element: 'pic',
            data: pic,
            nvPicPr: pic['nvPicPr'], 
            blipFill: pic['blipFill'],
            spPr: pic['spPr'],
            relationshipId: relationshipId
          });
        }
      }
    }
    
    return videos;
  }
  
  /**
   * Extract text from PPTX spTree
   */
  extractPPTXText(spTree: any): any[] {
    const textComponents = [];
    const spArray = this.ensureArray(spTree['sp']);
    
    for (const sp of spArray) {
      // Only process shapes with actual text content
      if (sp['txBody'] && this.hasTextContent(sp['txBody'])) {
        textComponents.push({
          type: 'text',
          namespace: 'p',
          element: 'sp',
          data: sp,
          spPr: sp['spPr'],
          nvSpPr: sp['nvSpPr'],
          style: sp['style'], // Add style data for consistency
          textBody: sp['txBody'] // Direct access
        });
      }
    }
    
    return textComponents;
  }
  
  /**
   * Extract text from clipboard lockedCanvas
   */
  extractClipboardText(lockedCanvas: any): any[] {
    const textComponents = [];
    const spData = lockedCanvas['sp'];
    
    if (spData) {
      const spArray = this.ensureArray(spData);
      
      for (const sp of spArray) {
        // Check for text content in the extra txSp layer
        if (sp['txSp'] && sp['txSp']['txBody']) {
          // Only treat as text component if there's actual text content
          if (this.hasTextContent(sp['txSp']['txBody'])) {
            textComponents.push({
              type: 'text',
              namespace: 'a',
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'], // Add style data for consistency
              textBody: sp['txSp']['txBody'] // Navigate extra layer
            });
          }
        }
      }
    }
    
    return textComponents;
  }
  
  /**
   * Extract media files from JSON
   */
  extractMediaFiles(json: any): Record<string, any> {
    const mediaFiles: Record<string, any> = {};
    const files = Object.keys(json);
    
    for (const file of files) {
      if (file.includes('/media/') && !file.endsWith('.xml')) {
        mediaFiles[file] = json[file];
      }
    }
    
    return mediaFiles;
  }
  
  /**
   * Extract relationship files from JSON
   */
  extractRelationships(json: any): Record<string, any> {
    const relationships: Record<string, any> = {};
    const files = Object.keys(json);
    
    for (const file of files) {
      if (file.includes('_rels/') && file.endsWith('.rels')) {
        relationships[file] = json[file];
      }
    }
    
    return relationships;
  }
  
  /**
   * Ensure value is an object (handle both array and object cases)
   */
  ensureObject(value: any): any | null {
    if (Array.isArray(value)) {
      return value[0] || null;
    }
    return value || null;
  }
  
  /**
   * Ensure value is an array
   */
  ensureArray(value: any): any[] {
    if (!value) return [];
    return Array.isArray(value) ? value : [value];
  }

  /**
   * Check if a textBody contains actual text content
   */
  hasTextContent(textBody: any): boolean {
    if (!textBody) return false;
    
    // Get paragraphs
    const paragraphs = textBody['p'];
    if (!paragraphs) return false;
    
    const paragraphArray = Array.isArray(paragraphs) ? paragraphs : [paragraphs];
    
    // Check each paragraph for actual text runs
    for (const p of paragraphArray) {
      const runs = p['r'];
      if (runs) {
        const runArray = Array.isArray(runs) ? runs : [runs];
        for (const run of runArray) {
          const text = run['t'];
          if (text && typeof text === 'string' && text.trim().length > 0) {
            return true;
          }
        }
      }
    }
    
    return false;
  }

  /**
   * Extract elements in their original z-order from spTree
   * This preserves the layering order that elements appear in PowerPoint
   */
  extractOrderedElements(spTree: any): NormalizedElement[] {
    const elements: NormalizedElement[] = [];
    let zIndex = 0;

    // Process all child elements in their original order
    // This preserves the back-to-front ordering from PowerPoint XML
    for (const [key, value] of Object.entries(spTree)) {
      if (key === 'sp') {
        // Handle shapes and text boxes
        const spArray = this.ensureArray(value);
        for (const sp of spArray) {
          if (sp['txBody'] && this.hasTextContent(sp['txBody'])) {
            // Text element
            elements.push({
              type: 'text',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txBody']
            });
          } else {
            // Shape element (non-text)
            // Skip text boxes (they're handled above)
            if (sp['nvSpPr'] && sp['nvSpPr']['cNvSpPr'] && sp['nvSpPr']['cNvSpPr']['$txBox']) continue;
            
            // Check if this shape uses background fill (design element)
            const useBgFill = sp['$useBgFill'] === '1' || sp['useBgFill'] === '1' || sp['$useBgFill'] === 1 || sp['useBgFill'] === 1;
            
            
            // Skip shapes that use background fill as they are typically design elements
            // that should not be rendered as regular shapes
            if (useBgFill) continue;
            
            elements.push({
              type: 'shape',
              zIndex: zIndex++,
              namespace: 'p', 
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txSp'] ? sp['txSp']['txBody'] : null
            });
          }
        }
      } else if (key === 'pic') {
        // Handle images and videos
        const picArray = this.ensureArray(value);
        for (const pic of picArray) {
          // Check if this is a video by looking for nvPr.videoFile
          const nvPicPr = pic['nvPicPr'];
          const isVideo = nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile'];
          
          if (isVideo) {
            // Video element
            const videoFile = nvPicPr['nvPr']['videoFile'];
            const relationshipId = videoFile['$link'] || videoFile['link'];
            
            elements.push({
              type: 'video',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr'],
              relationshipId: relationshipId
            });
          } else {
            // Image element
            elements.push({
              type: 'image',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr']
            });
          }
        }
      } else if (key === 'graphicFrame') {
        // Handle tables (graphicFrame containing table data)
        const graphicFrameArray = this.ensureArray(value);
        for (const graphicFrame of graphicFrameArray) {
          // Check if this is a table by looking at the graphicData URI
          const graphic = graphicFrame['graphic'];
          const graphicData = graphic?.['graphicData'];
          const uri = graphicData?.['$uri'];
          
          if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            elements.push({
              type: 'table',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'graphicFrame',
              data: graphicFrame,
              nvGraphicFramePr: graphicFrame['nvGraphicFramePr'],
              spPr: graphicFrame['xfrm'], // Use xfrm for positioning
              graphicData: graphicData
            });
          } else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/diagram') {
            elements.push({
              type: 'diagram',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'graphicFrame',
              data: graphicFrame,
              nvGraphicFramePr: graphicFrame['nvGraphicFramePr'],
              spPr: graphicFrame['xfrm'], // Use xfrm for positioning
              graphicData: graphicData
            });
          }
        }
      } else if (key === 'cxnSp') {
        // Handle connection shapes (arrows, connectors)
        const cxnSpArray = this.ensureArray(value);
        for (const cxnSp of cxnSpArray) {
          elements.push({
            type: 'connection',
            zIndex: zIndex++,
            namespace: 'p',
            element: 'cxnSp',
            data: cxnSp,
            nvCxnSpPr: cxnSp['nvCxnSpPr'],
            spPr: cxnSp['spPr'],
            startConnection: cxnSp['nvCxnSpPr']?.['cNvCxnSpPr']?.['stCxn'],
            endConnection: cxnSp['nvCxnSpPr']?.['cNvCxnSpPr']?.['endCxn']
          });
        }
      } else if (key === 'grpSp') {
        // Handle grouped shapes - flatten all child elements
        const grpSpArray = this.ensureArray(value);
        for (const grpSp of grpSpArray) {
          // Recursively extract elements from the group
          const groupElements = this.extractElementsFromGroup(grpSp, zIndex);
          elements.push(...groupElements);
          zIndex += groupElements.length;
        }
      }
    }

    return elements;
  }

  /**
   * Extract elements from a grouped shape (grpSp), flattening the group structure
   */
  extractElementsFromGroup(grpSp: any, startZIndex: number): NormalizedElement[] {
    const elements: NormalizedElement[] = [];
    let zIndex = startZIndex;

    // Process all child elements within the group
    for (const [key, value] of Object.entries(grpSp)) {
      if (key === 'sp') {
        // Handle shapes and text boxes within the group
        const spArray = this.ensureArray(value);
        for (const sp of spArray) {
          if (sp['txBody'] && this.hasTextContent(sp['txBody'])) {
            // Text element
            elements.push({
              type: 'text',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txBody']
            });
          } else {
            // Shape element (non-text)
            // Skip text boxes (they're handled above)
            if (sp['nvSpPr'] && sp['nvSpPr']['cNvSpPr'] && sp['nvSpPr']['cNvSpPr']['$txBox']) continue;
            
            elements.push({
              type: 'shape',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txSp'] ? sp['txSp']['txBody'] : null
            });
          }
        }
      } else if (key === 'pic') {
        // Handle images within the group
        const picArray = this.ensureArray(value);
        for (const pic of picArray) {
          // Check if this is a video by looking for nvPr.videoFile
          const nvPicPr = pic['nvPicPr'];
          const isVideo = nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile'];
          
          if (isVideo) {
            // Video element
            const videoFile = nvPicPr['nvPr']['videoFile'];
            const relationshipId = videoFile['$link'] || videoFile['link'];
            
            elements.push({
              type: 'video',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr'],
              relationshipId: relationshipId
            });
          } else {
            // Image element
            elements.push({
              type: 'image',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr']
            });
          }
        }
      } else if (key === 'graphicFrame') {
        // Handle tables within the group
        const graphicFrameArray = this.ensureArray(value);
        for (const graphicFrame of graphicFrameArray) {
          // Check if this is a table by looking at the graphicData URI
          const graphic = graphicFrame['graphic'];
          const graphicData = graphic?.['graphicData'];
          const uri = graphicData?.['$uri'];
          
          if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            elements.push({
              type: 'table',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'graphicFrame',
              data: graphicFrame,
              nvGraphicFramePr: graphicFrame['nvGraphicFramePr'],
              spPr: graphicFrame['xfrm'], // Use xfrm for positioning
              graphicData: graphicData
            });
          } else if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/diagram') {
            elements.push({
              type: 'diagram',
              zIndex: zIndex++,
              namespace: 'p',
              element: 'graphicFrame',
              data: graphicFrame,
              nvGraphicFramePr: graphicFrame['nvGraphicFramePr'],
              spPr: graphicFrame['xfrm'], // Use xfrm for positioning
              graphicData: graphicData
            });
          }
        }
      } else if (key === 'cxnSp') {
        // Handle connection shapes within groups (arrows, connectors)
        const cxnSpArray = this.ensureArray(value);
        for (const cxnSp of cxnSpArray) {
          elements.push({
            type: 'connection',
            zIndex: zIndex++,
            namespace: 'p',
            element: 'cxnSp',
            data: cxnSp,
            nvCxnSpPr: cxnSp['nvCxnSpPr'],
            spPr: cxnSp['spPr'],
            startConnection: cxnSp['nvCxnSpPr']?.['cNvCxnSpPr']?.['stCxn'],
            endConnection: cxnSp['nvCxnSpPr']?.['cNvCxnSpPr']?.['endCxn']
          });
        }
      } else if (key === 'grpSp') {
        // Handle nested groups (recursive)
        const nestedGrpSpArray = this.ensureArray(value);
        for (const nestedGrpSp of nestedGrpSpArray) {
          const nestedElements = this.extractElementsFromGroup(nestedGrpSp, zIndex);
          elements.push(...nestedElements);
          zIndex += nestedElements.length;
        }
      }
    }

    return elements;
  }

  /**
   * Extract elements in their original z-order from clipboard lockedCanvas
   * This preserves the layering order that elements appear in PowerPoint clipboard
   */
  extractOrderedClipboardElements(lockedCanvas: any): NormalizedElement[] {
    const elements: NormalizedElement[] = [];
    let zIndex = 0;

    // Process all child elements in their original order
    // This preserves the back-to-front ordering from PowerPoint clipboard XML
    for (const [key, value] of Object.entries(lockedCanvas)) {
      if (key === 'sp') {
        // Handle shapes and text boxes
        const spArray = this.ensureArray(value);
        for (const sp of spArray) {
          if (sp['txSp'] && sp['txSp']['txBody'] && this.hasTextContent(sp['txSp']['txBody'])) {
            // Text element with clipboard structure
            elements.push({
              type: 'text',
              zIndex: zIndex++,
              namespace: 'a',
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txSp']['txBody'] // Navigate extra layer
            });
          } else {
            // Shape element (non-text)
            elements.push({
              type: 'shape',
              zIndex: zIndex++,
              namespace: 'a', 
              element: 'sp',
              data: sp,
              spPr: sp['spPr'],
              nvSpPr: sp['nvSpPr'],
              style: sp['style'],
              textBody: sp['txSp'] ? sp['txSp']['txBody'] : null
            });
          }
        }
      } else if (key === 'pic') {
        // Handle images and videos
        const picArray = this.ensureArray(value);
        for (const pic of picArray) {
          // Check if this is a video by looking for nvPr.videoFile
          const nvPicPr = pic['nvPicPr'];
          const isVideo = nvPicPr && nvPicPr['nvPr'] && nvPicPr['nvPr']['videoFile'];
          
          if (isVideo) {
            // Video element
            const videoFile = nvPicPr['nvPr']['videoFile'];
            const relationshipId = videoFile['$link'] || videoFile['link'];
            
            elements.push({
              type: 'video',
              zIndex: zIndex++,
              namespace: 'a',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr'],
              relationshipId: relationshipId
            });
          } else {
            // Image element
            elements.push({
              type: 'image',
              zIndex: zIndex++,
              namespace: 'a',
              element: 'pic',
              data: pic,
              nvPicPr: pic['nvPicPr'],
              blipFill: pic['blipFill'],
              spPr: pic['spPr']
            });
          }
        }
      } else if (key === 'graphicFrame') {
        // Handle tables (graphicFrame containing table data) in clipboard
        const graphicFrameArray = this.ensureArray(value);
        for (const graphicFrame of graphicFrameArray) {
          // Check if this is a table by looking at the graphicData URI
          const graphic = graphicFrame['graphic'];
          const graphicData = graphic?.['graphicData'];
          const uri = graphicData?.['$uri'];
          
          if (uri === 'http://schemas.openxmlformats.org/drawingml/2006/table') {
            elements.push({
              type: 'table',
              zIndex: zIndex++,
              namespace: 'a',
              element: 'graphicFrame',
              data: graphicFrame,
              nvGraphicFramePr: graphicFrame['nvGraphicFramePr'],
              spPr: graphicFrame['xfrm'], // Use xfrm for positioning
              graphicData: graphicData
            });
          }
        }
      }
    }

    return elements;
  }

  /**
   * Recursively strip namespace prefixes from all object keys
   * Converts 'p:spPr' -> 'spPr', 'a:xfrm' -> 'xfrm', etc.
   */
  stripNamespaces(obj: any): any {
    if (obj === null || obj === undefined) {
      return obj;
    }

    // Handle binary data (Uint8Array, ArrayBuffer, Buffer) - don't process recursively
    if (obj instanceof Uint8Array || obj instanceof ArrayBuffer || 
        (typeof Buffer !== 'undefined' && Buffer.isBuffer && Buffer.isBuffer(obj))) {
      return obj;
    }

    if (Array.isArray(obj)) {
      return obj.map(item => this.stripNamespaces(item));
    }

    if (typeof obj === 'object') {
      const stripped: any = {};
      
      for (const [key, value] of Object.entries(obj)) {
        // Strip namespace prefix (everything before and including the colon)
        const strippedKey = key.includes(':') ? key.split(':')[1] : key;
        stripped[strippedKey] = this.stripNamespaces(value);
      }
      
      return stripped;
    }

    // Primitive values (string, number, boolean) pass through unchanged
    return obj;
  }

  /**
   * Convert layout element to slide format for consistent processing
   */
  convertLayoutElementToSlideFormat(layoutElement: any, slideSpTree: any): NormalizedElement | null {
    if (!layoutElement || !layoutElement.type || !layoutElement.data) {
      return null;
    }

    try {
      const { type, data, zIndex, isLayoutElement, isBackgroundElement } = layoutElement;

      if (type === 'shape') {
        // Only convert shapes that have required properties for parsing
        if (!data['spPr']) {
          console.warn('Layout shape missing spPr, skipping');
          return null;
        }
        
        // Convert shape to normalized format
        return {
          type: 'shape',
          zIndex: zIndex,
          namespace: 'p',
          element: 'sp',
          data: data,
          spPr: data['spPr'],
          nvSpPr: data['nvSpPr'],
          style: data['style'],
          isLayoutElement: true,
          isBackgroundElement: isBackgroundElement || false
        };
      } else if (type === 'image') {
        // Only convert images that have required properties
        if (!data['spPr'] && !data['nvPicPr']) {
          console.warn('Layout image missing required properties, skipping');
          return null;
        }
        
        // Convert image to normalized format
        return {
          type: 'image',
          zIndex: zIndex,
          namespace: 'p',
          element: 'pic',
          data: data,
          nvPicPr: data['nvPicPr'],
          blipFill: data['blipFill'],
          spPr: data['spPr'],
          isLayoutElement: true,
          isBackgroundElement: isBackgroundElement || false,
          relationshipId: data.relationshipId
        };
      } else if (type === 'text') {
        // Only convert text that has required properties
        if (!data['spPr']) {
          console.warn('Layout text missing spPr, skipping');
          return null;
        }
        
        // Convert text to normalized format
        return {
          type: 'text',
          zIndex: zIndex,
          namespace: 'p',
          element: 'sp',
          data: data,
          spPr: data['spPr'],
          nvSpPr: data['nvSpPr'],
          txBody: data['txBody'],
          isLayoutElement: true
        };
      }

      return null;
    } catch (error) {
      console.warn('Error converting layout element:', error);
      return null;
    }
  }

  /**
   * Convert master element to slide format for consistent processing
   */
  convertMasterElementToSlideFormat(masterElement: any, slideSpTree: any): NormalizedElement | null {
    if (!masterElement || !masterElement.type || !masterElement.data) {
      return null;
    }

    try {
      const { type, data, zIndex, isMasterElement, isBackgroundElement } = masterElement;

      if (type === 'shape') {
        // Only convert shapes that have required properties for parsing
        if (!data['spPr']) {
          console.warn('Master shape missing spPr, skipping');
          return null;
        }
        
        // Convert shape to normalized format
        return {
          type: 'shape',
          zIndex: zIndex,
          namespace: 'p',
          element: 'sp',
          data: data,
          spPr: data['spPr'],
          nvSpPr: data['nvSpPr'],
          style: data['style'],
          isMasterElement: true,
          isBackgroundElement: isBackgroundElement || false
        };
      } else if (type === 'image') {
        // Only convert images that have required properties
        if (!data['spPr'] && !data['nvPicPr']) {
          console.warn('Master image missing required properties, skipping');
          return null;
        }
        
        // Convert image to normalized format
        return {
          type: 'image',
          zIndex: zIndex,
          namespace: 'p',
          element: 'pic',
          data: data,
          nvPicPr: data['nvPicPr'],
          blipFill: data['blipFill'],
          spPr: data['spPr'],
          isMasterElement: true,
          isBackgroundElement: isBackgroundElement || false,
          relationshipId: data.relationshipId
        };
      } else if (type === 'text') {
        // Only convert text that has required properties
        if (!data['spPr']) {
          console.warn('Master text missing spPr, skipping');
          return null;
        }
        
        // Convert text to normalized format
        return {
          type: 'text',
          zIndex: zIndex,
          namespace: 'p',
          element: 'sp',
          data: data,
          spPr: data['spPr'],
          nvSpPr: data['nvSpPr'],
          txBody: data['txBody'],
          isMasterElement: true
        };
      }

      return null;
    } catch (error) {
      console.warn('Error converting master element:', error);
      return null;
    }
  }

  /**
   * Calculate content bounds to determine appropriate slide dimensions
   */
  calculateContentBounds(slides: NormalizedSlide[]): { width: number; height: number } {
    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    let hasContent = false;

    // BaseParser is already imported at the top

    for (const slide of slides) {
      // Check all element types
      const allElements = [...slide.shapes, ...slide.text, ...slide.images];
      
      for (const element of allElements) {
        if (!element.data || !element.spPr) continue;
        
        try {
          // Parse transform to get position and size
          const xfrm = element.spPr['xfrm'];
          if (!xfrm) continue;
          
          const transform = BaseParser.parseTransform(xfrm);
          if (transform.width === 0 && transform.height === 0) continue;
          
          // Calculate element bounds
          const elementMinX = transform.x;
          const elementMinY = transform.y;
          const elementMaxX = transform.x + transform.width;
          const elementMaxY = transform.y + transform.height;
          
          // Update overall bounds
          minX = Math.min(minX, elementMinX);
          minY = Math.min(minY, elementMinY);
          maxX = Math.max(maxX, elementMaxX);
          maxY = Math.max(maxY, elementMaxY);
          
          hasContent = true;
        } catch (error) {
          console.warn('Error parsing element bounds:', error);
        }
      }
    }
    
    if (!hasContent) {
      // Fallback to default dimensions if no content found
      // console.log('No content bounds found, using default dimensions');
      return { width: DEFAULT_SLIDE_WIDTH_PX, height: DEFAULT_SLIDE_HEIGHT_PX };
    }
    
    // Add padding around content (10% on each side)
    const contentWidth = maxX - minX;
    const contentHeight = maxY - minY;
    const paddingX = Math.max(50, contentWidth * 0.1); // Minimum 50px padding
    const paddingY = Math.max(50, contentHeight * 0.1); // Minimum 50px padding
    
    const finalWidth = Math.round(contentWidth + (paddingX * 2));
    const finalHeight = Math.round(contentHeight + (paddingY * 2));
    
    // Ensure minimum dimensions
    const minWidth = 400;
    const minHeight = 300;
    
    const result = {
      width: Math.max(finalWidth, minWidth),
      height: Math.max(finalHeight, minHeight)
    };
    
    // Calculated content-based slide dimensions
    
    return result;
  }

  /**
   * Extract theme data from PowerPoint files
   * @param json The parsed PowerPoint JSON
   * @param format The detected format (pptx or clipboard)
   * @returns Parsed theme data or null
   */
  extractThemeData(json: any, format: string): any | null {
    try {
      // Determine theme file path based on format
      const themeFilePath = format === 'pptx' ? 'ppt/theme/theme1.xml' : 'clipboard/theme/theme1.xml';
      
      const themeFile = json[themeFilePath];
      if (!themeFile) {
        // console.log(`No theme file found at ${themeFilePath}`);
        return null;
      }

      // Extract theme from the XML structure
      const theme = themeFile.theme;
      if (!theme) {
        // console.log('No theme element found in theme file');
        return null;
      }

      // Parse the theme color scheme
      const themeElements = theme.themeElements;
      if (!themeElements) {
        // console.log('No themeElements found in theme');
        return null;
      }

      const clrScheme = themeElements.clrScheme;
      if (!clrScheme) {
        // console.log('No clrScheme found in themeElements');
        return null;
      }

      // Extract actual theme colors
      const themeColors: Record<string, string> = {};
      
      // Extract each color scheme color
      const colorNames = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      
      for (const colorName of colorNames) {
        const colorDef = clrScheme[colorName];
        if (colorDef) {
          const parsedColor = this.parseThemeColor(colorDef);
          if (parsedColor) {
            themeColors[colorName] = parsedColor;
          }
        }
      }

      // console.log('Extracted theme colors:', Object.keys(themeColors).length, 'colors');
      return {
        colors: themeColors,
        rawTheme: theme // Keep raw theme for advanced processing
      };
    } catch (error) {
      console.warn('Error extracting theme data:', error);
      return null;
    }
  }

  /**
   * Parse a theme color definition to hex color
   * @param colorDef Theme color definition from XML
   * @returns Hex color string or null
   */
  private parseThemeColor(colorDef: any): string | null {
    try {
      // Handle srgbClr (direct RGB)
      if (colorDef.srgbClr) {
        const val = colorDef.srgbClr.$val;
        return val && typeof val === 'string' ? `#${val.toUpperCase()}` : null;
      }

      // Handle sysClr (system colors)
      if (colorDef.sysClr) {
        const val = colorDef.sysClr.$val;
        const lastClr = colorDef.sysClr.$lastClr;
        // Prefer lastClr if available, fallback to system color name
        if (lastClr && typeof lastClr === 'string') {
          return `#${lastClr.toUpperCase()}`;
        }
        // Map common system colors
        const sysColors: Record<string, string> = {
          'windowText': '#000000',
          'window': '#FFFFFF',
          'btnText': '#000000',
          'btnFace': '#F0F0F0'
        };
        return sysColors[val] || '#000000';
      }

      console.warn('Unknown theme color format:', colorDef);
      return null;
    } catch (error) {
      console.warn('Error parsing theme color:', error);
      return null;
    }
  }
}