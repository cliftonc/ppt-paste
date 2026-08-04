/**
 * Main PowerPoint parser that coordinates all specialized parsers
 */

import { PowerPointNormalizer } from './PowerPointNormalizer.js';
import { TextParser } from './TextParser.js';
import { ShapeParser } from './ShapeParser.js';
import { ImageParser } from './ImageParser.js';
import { TableParser } from './TableParser.js';
import { DiagramParser } from './DiagramParser.js';
import { VideoParser } from './VideoParser.js';
import { ConnectorParser } from './ConnectorParser.js';
import { BaseParser } from './BaseParser.js';
import type { PowerPointComponent, ConnectionComponent, PlaceholderMap, XMLNode } from '../types/index.js';

import { isTextElement, isShapeElement, isImageElement, isTableElement, isDiagramElement, isVideoElement, isConnectionElement, type NormalizedTextElement, type NormalizedShapeElement, type NormalizedImageElement, type NormalizedTableElement, type NormalizedDiagramElement, type NormalizedVideoElement, type NormalizedConnectionElement, type MediaFiles, type RelationshipGraph, type NormalizedSlide } from '../types/normalized.js';

export interface R2BucketLike {
  put?(key: string, value: any, options?: any): Promise<any> | any;
  get?(key: string): Promise<any> | any;
  head?(key: string): Promise<any> | any;
}

export interface ParseOptions {
  debug?: boolean;
  r2Storage?: R2BucketLike | null;
}

export interface SlideMetadata {
  name: string;
  componentCount: number;
  format: string;
  slideFile: string | null;
  layoutFile: string | null;
  masterFile: string | null;
  layoutElementCount: number;
  masterElementCount: number;
  /** Notes part resolved through the rels graph, or null when the slide has none */
  notesFile: string | null;
  /** The number in the slide's filename — creation order, not presentation order */
  fileSlideNumber: number | null;
  /** How the title was identified: by the slide's own placeholder, or via the layout */
  titleSource: 'placeholder' | 'layout' | null;
}

export interface ParsedMaster {
  id: string;
  name: string;
  background?: PowerPointComponent;
  components: PowerPointComponent[];
  sourceFile: string;
  placeholders?: PlaceholderMap;
  textStyles?: { titleStyle?: XMLNode, bodyStyle?: XMLNode };
}

export interface ParsedLayout {
  id: string;
  name: string;
  masterId?: string;
  background?: PowerPointComponent;
  components: PowerPointComponent[];
  sourceFile: string;
  placeholders?: PlaceholderMap;
}

export interface ParsedSlide {
  slideIndex: number; // zero-based position in presentation order
  slideNumber: number; // 1-based position in presentation order (what the audience sees)
  title?: string; // Text of the title placeholder, when the slide has one
  notes?: string; // Speaker notes, when the slide has any
  layoutId?: string; // Reference to layout
  background?: PowerPointComponent; // Slide-specific background
  components: PowerPointComponent[];
  metadata: SlideMetadata;
}

export interface ParsedResult {
  slides: ParsedSlide[];
  masters: Record<string, ParsedMaster>;
  layouts: Record<string, ParsedLayout>;
  totalComponents: number;
  format: string; // 'pptx' | 'clipboard'
  /** 'presentation' when p:sldIdLst decided the order, 'filename' when it could not be read */
  slideOrderSource?: 'presentation' | 'filename';
  slideDimensions?: {
    width: number;
    height: number;
  };
}

export class PowerPointParser extends BaseParser {
  private normalizer: PowerPointNormalizer;
  private json: any;
  private normalizedJson: any;

  constructor() {
    super();
    this.normalizer = new PowerPointNormalizer();
  }

  /**
   * Parse PowerPoint JSON data into structured slides
   */
  async parseJson(json: any, options: ParseOptions = {}): Promise<ParsedResult> {
    const { debug = false, r2Storage = null } = options;
    
    // Store both raw and normalized json for access in component parsers
    this.json = json;
    // Strip namespaces from the entire JSON structure for DiagramParser access
    this.normalizedJson = this.normalizer.stripNamespaces(json);
    
    try {
      // if (debug) console.log('Processing PowerPoint JSON data...');
      
      // Step 1: Normalize the structure (eliminates all format differences!)
      const normalized = this.normalizer.normalize(json);
      if (debug) {
        // console.log(`Normalized ${normalized.format} format with ${normalized.slides.length} slides`);
        if (normalized.theme?.colors) {
          // console.log(`Theme colors available:`, Object.keys(normalized.theme.colors));
        }
      }
      
      // Set theme colors for this parsing session
      BaseParser.setThemeColors(normalized.theme?.colors);
      
      // Step 2: Extract masters and layouts first
      const masters: Record<string, ParsedMaster> = {};
      const layouts: Record<string, ParsedLayout> = {};
      let globalComponentIndex = 0;
      
      
      // Process masters and layouts directly from available files
      const masterFiles = new Set<string>();
      const layoutFiles = new Set<string>();
      
      // Find all master and layout files directly from JSON
      const allFiles = Object.keys(json);
      allFiles.forEach(filePath => {
        if (filePath.startsWith('ppt/slideMasters/') && filePath.endsWith('.xml')) {
          masterFiles.add(filePath);
        }
        if (filePath.startsWith('ppt/slideLayouts/') && filePath.endsWith('.xml')) {
          layoutFiles.add(filePath);
        }
      });
      
      
      // Extract master definitions
      if (normalized.format === 'pptx' && json) {
        const pptxParser = new (await import('../processors/PPTXParser.js')).PPTXParser();
        
        // Use normalized relationships and media files for consistency
        const relationships = normalized.relationships;
        const mediaFiles = normalized.mediaFiles;
        
        // Process masters and layouts from the normalized data
        // This way we reuse all the existing slide parsing logic
        const masterSlides = new Map<string, NormalizedSlide>();
        const layoutSlides = new Map<string, NormalizedSlide>();
        
        // Process all available layout files directly using PPTXParser
        for (const layoutFile of layoutFiles) {
          try {
            const layoutElements = pptxParser.getSlideLayoutElements(json, layoutFile, normalized.slideDimensions);
            // Always process the layout, even if it has no elements, because it might have placeholders
            const layoutSlide: NormalizedSlide = {
              slideFile: layoutFile,
              format: 'pptx',
              shapes: [],
              images: [],
              text: [],
              videos: [],
              elements: layoutElements.map(el => ({
                type: el.type as any,
                zIndex: el.zIndex,
                namespace: 'p',
                element: 'sp',
                data: el.data,
                isLayoutElement: true
              })),
              layoutFile,
              masterFile: null
            };
            layoutSlides.set(layoutFile, layoutSlide);
          } catch (error) {
            console.warn(`Failed to process layout ${layoutFile}:`, error);
          }
        }
        
        // Process all available master files directly using PPTXParser
        for (const masterFile of masterFiles) {
          try {
            const masterElements = pptxParser.getSlideMasterElements(json, masterFile, normalized.slideDimensions);
            if (masterElements.length > 0) {
              // Create a synthetic normalized slide for this master
              const masterSlide: NormalizedSlide = {
                slideFile: masterFile,
                format: 'pptx',
                shapes: [],
                images: [],
                text: [],
                videos: [],
                elements: masterElements.map(el => ({
                  type: el.type as any,
                  zIndex: el.zIndex,
                  namespace: 'p',
                  element: 'sp',
                  data: el.data,
                  isMasterElement: true
                })),
                layoutFile: null,
                masterFile
              };
              masterSlides.set(masterFile, masterSlide);
            }
          } catch (error) {
            console.warn(`Failed to process master ${masterFile}:`, error);
          }
        }

        // Process master definitions using loaded data
        for (const [masterFile, masterSlide] of masterSlides) {
          const masterId = masterFile.replace('ppt/slideMasters/', '').replace('.xml', '');
          const masterComponents: PowerPointComponent[] = [];
          let masterBackground: PowerPointComponent | undefined;
          
          // Process master elements using existing slide logic
          for (let j = 0; j < masterSlide.elements.length; j++) {
            const element = masterSlide.elements[j];
            
            if (!element.isMasterElement) continue; // Only process master elements
            
            if (element.isBackgroundElement) {
              // if (debug) console.log(`Processing master background element`);
              if (isShapeElement(element)) {
                const bgComponent = await this.parseUnifiedShapeComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, -2000, { debug, isMasterOrLayout: true });
                // Only set as background if it's not a white/near-white background
                // Exception: Always allow image backgrounds regardless of fill color
                if (bgComponent && (bgComponent.type === 'image' || (bgComponent.style?.fillColor && !BaseParser.isWhiteOrNearWhite(bgComponent.style.fillColor)))) {
                  masterBackground = bgComponent;
                }
              }
              continue;
            }
            
            let component: PowerPointComponent | null = null;
            
            if (isTextElement(element)) {
              component = await this.parseUnifiedTextComponent(element, globalComponentIndex++, 0, element.zIndex, { debug });
            } else if (isShapeElement(element)) {
              component = await this.parseUnifiedShapeComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, isMasterOrLayout: true });
            } else if (isImageElement(element)) {
              component = await this.parseUnifiedImageComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, r2Storage });
            } else if (isTableElement(element)) {
              component = await this.parseUnifiedTableComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug });
            } else if (isDiagramElement(element)) {
              const diagramResult = await this.parseUnifiedDiagramComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug });
              if (Array.isArray(diagramResult)) {
                masterComponents.push(...diagramResult);
                continue;
              } else {
                component = diagramResult;
              }
            } else if (isVideoElement(element)) {
              component = await this.parseUnifiedVideoComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, r2Storage });
            } else if (isConnectionElement(element)) {
              component = await this.parseUnifiedConnectionComponent(element, globalComponentIndex++, 0, element.zIndex, { debug });
            }
            
            if (component) {
              masterComponents.push(component);
            }
          }
          
          // Extract placeholder definitions from master
          const masterPlaceholders = pptxParser.getPlaceholderDefinitions(json, masterFile);
          
          // Extract text styles (txStyles) from master
          const masterStyles = this.extractMasterTextStyles(json, masterFile);
          
          masters[masterFile] = {
            id: masterId,
            name: `Master ${masterId}`,
            background: masterBackground,
            components: masterComponents,
            sourceFile: masterFile,
            placeholders: masterPlaceholders,
            textStyles: masterStyles
          };
          
          if (debug) {
            // console.log(`Extracted master ${masterId}: ${masterBackground ? '1 background' : '0 backgrounds'}, ${masterComponents.length} components`);
          }
        }
        
        // Process layout definitions using normalized data
        for (const [layoutFile, layoutSlide] of layoutSlides) {
          const layoutId = layoutFile.replace('ppt/slideLayouts/', '').replace('.xml', '');
          const layoutComponents: PowerPointComponent[] = [];
          let layoutBackground: PowerPointComponent | undefined;
          let masterId: string | undefined;
          
          // Find which master this layout uses
          const layoutMasterRels = pptxParser.getLayoutMasterRelationships(json);
          const masterFile = layoutMasterRels[layoutFile];
          if (masterFile) {
            masterId = masterFile.replace('ppt/slideMasters/', '').replace('.xml', '');
          }
          
          // Process layout elements using existing slide logic
          for (let j = 0; j < layoutSlide.elements.length; j++) {
            const element = layoutSlide.elements[j];
            
            if (!element.isLayoutElement) continue; // Only process layout elements
            
            if (element.isBackgroundElement) {
              // if (debug) console.log(`Processing layout background element`);
              if (isShapeElement(element)) {
                const bgComponent = await this.parseUnifiedShapeComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, -1000, { debug, isMasterOrLayout: true });
                // Only set as background if it's not a white/near-white background
                // Exception: Always allow image backgrounds regardless of fill color
                if (bgComponent && (bgComponent.type === 'image' || (bgComponent.style?.fillColor && !BaseParser.isWhiteOrNearWhite(bgComponent.style.fillColor)))) {
                  layoutBackground = bgComponent;
                }
              }
              continue;
            }
            
            let component: PowerPointComponent | null = null;
            
            if (isTextElement(element)) {
              component = await this.parseUnifiedTextComponent(element, globalComponentIndex++, 0, element.zIndex, { debug });
            } else if (isShapeElement(element)) {
              component = await this.parseUnifiedShapeComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, isMasterOrLayout: true });
            } else if (isImageElement(element)) {
              component = await this.parseUnifiedImageComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, r2Storage });
            } else if (isTableElement(element)) {
              component = await this.parseUnifiedTableComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug });
            } else if (isDiagramElement(element)) {
              const diagramResult = await this.parseUnifiedDiagramComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug });
              if (Array.isArray(diagramResult)) {
                layoutComponents.push(...diagramResult);
                continue;
              } else {
                component = diagramResult;
              }
            } else if (isVideoElement(element)) {
              component = await this.parseUnifiedVideoComponent(element, relationships, mediaFiles, globalComponentIndex++, 0, element.zIndex, { debug, r2Storage });
            } else if (isConnectionElement(element)) {
              component = await this.parseUnifiedConnectionComponent(element, globalComponentIndex++, 0, element.zIndex, { debug });
            }
            
            if (component) {
              layoutComponents.push(component);
            }
          }
          
          // Extract placeholder definitions from layout
          const layoutPlaceholders = pptxParser.getPlaceholderDefinitions(json, layoutFile);
          
          layouts[layoutFile] = {
            id: layoutId,
            name: `Layout ${layoutId}`,
            masterId,
            background: layoutBackground,
            components: layoutComponents,
            sourceFile: layoutFile,
            placeholders: layoutPlaceholders
          };
          
          if (debug) {
            // console.log(`Extracted layout ${layoutId}: ${layoutBackground ? '1 background' : '0 backgrounds'}, ${layoutComponents.length} components`);
          }
        }
      }
      
      // Step 3: Process all slides using unified structure
      const slides: ParsedSlide[] = [];
      const components: PowerPointComponent[] = []; // Keep flat array for backward compatibility
      
      for (let arrayIndex = 0; arrayIndex < normalized.slides.length; arrayIndex++) {
        const slide = normalized.slides[arrayIndex];
        const slideNumber = slide.slideNumber || (arrayIndex + 1); // Use extracted number or fallback
        
        if (debug) {
          // console.log(`Processing slide ${slideNumber}: ${slide.shapes.length} shapes, ${slide.text.length} text, ${slide.images.length} images`);
        }
        
        const slideComponents: PowerPointComponent[] = [];
        let slideBackground: PowerPointComponent | undefined;
        const layoutId = slide.layoutFile ? slide.layoutFile.replace('ppt/slideLayouts/', '').replace('.xml', '') : undefined;
        let localComponentIndex = 0;
        // Media, videos and SmartArt parts are resolved through THIS slide's rels part.
        // It is passed explicitly because the alternative — `ppt/slides/_rels/slide${slideNumber}.xml.rels`
        // — silently reads another slide's relationships once a deck has been reordered.
        const slideRelsFile = slide.slideRelsFile ?? null;
        
        // Build combined placeholder map for this slide (master → layout inheritance)
        let combinedPlaceholders: PlaceholderMap = {};
        let masterStyles: { titleStyle?: XMLNode, bodyStyle?: XMLNode } | undefined;
        if (slide.layoutFile && layouts[slide.layoutFile]) {
          const layout = layouts[slide.layoutFile];
          // Start with master placeholders if layout has a master
          if (layout.masterId && slide.masterFile && masters[slide.masterFile]) {
            combinedPlaceholders = { ...masters[slide.masterFile].placeholders || {} };
            // Extract master text styles for font inheritance
            masterStyles = masters[slide.masterFile].textStyles;
          }
          // Override with layout-specific placeholders
          if (layout.placeholders) {
            combinedPlaceholders = { ...combinedPlaceholders, ...layout.placeholders };
          }
        }
        
        // Note: Slide backgrounds are already extracted during normalization as elements with isBackgroundElement=true
        
        // Process components in their original z-order if available
        if (slide.elements && slide.elements.length > 0) {
          // Use ordered elements to preserve z-index
          for (const element of slide.elements) {
            // Skip master and layout backgrounds - they're now in separate definitions
            if (element.isBackgroundElement && (element.isMasterElement || element.isLayoutElement)) {
              continue;
            }
            
            // Handle slide-specific background
            if (element.isBackgroundElement && !element.isMasterElement && !element.isLayoutElement) {
              const bgComponent = await this.parseElementToComponent(element, normalized.relationships, normalized.mediaFiles, globalComponentIndex++, slideNumber - 1, -500, { debug, r2Storage, slideRelsFile });
              // Only set as background if it's not a white/near-white background
              // Exception: Always allow image backgrounds regardless of fill color
              if (bgComponent && (bgComponent.type === 'image' || (bgComponent.style?.fillColor && !BaseParser.isWhiteOrNearWhite(bgComponent.style.fillColor)))) {
                slideBackground = bgComponent;
                slideComponents.push(bgComponent);
                components.push(bgComponent);
              }
              continue;
            }
            
            let component: PowerPointComponent | null = null;
            
            if (isTextElement(element)) {
              component = await this.parseUnifiedTextComponent(
                element,
                globalComponentIndex++,
                slideNumber - 1,
                element.zIndex,
                { debug, placeholders: combinedPlaceholders, masterStyles }
              );
            } else if (isShapeElement(element)) {
              component = await this.parseUnifiedShapeComponent(
                element,
                normalized.relationships,
                normalized.mediaFiles,
                globalComponentIndex++,
                slideNumber - 1, // relationships index
                element.zIndex,
                { debug, placeholders: combinedPlaceholders }
              );
            } else if (isImageElement(element)) {
              component = await this.parseUnifiedImageComponent(
                element,
                normalized.relationships,
                normalized.mediaFiles,
                globalComponentIndex++,
                slideNumber - 1, // relationships index
                element.zIndex,
                { debug, r2Storage, slideRelsFile }
              );
            } else if (isTableElement(element)) {
              component = await this.parseUnifiedTableComponent(
                element,
                normalized.relationships,
                normalized.mediaFiles,
                globalComponentIndex++,
                slideNumber - 1,
                element.zIndex,
                { debug }
              );
            } else if (isDiagramElement(element)) {
              const diagramResult = await this.parseUnifiedDiagramComponent(
                element,
                normalized.relationships,
                normalized.mediaFiles,
                globalComponentIndex++,
                slideNumber - 1,
                element.zIndex,
                { debug, slideRelsFile }
              );
              if (Array.isArray(diagramResult)) {
                slideComponents.push(...diagramResult);
                continue;
              } else {
                component = diagramResult;
              }
            } else if (isVideoElement(element)) {
              component = await this.parseUnifiedVideoComponent(
                element,
                normalized.relationships,
                normalized.mediaFiles,
                globalComponentIndex++,
                slideNumber - 1,
                element.zIndex,
                { debug, r2Storage, slideRelsFile }
              );
            } else if (isConnectionElement(element)) {
              component = await this.parseUnifiedConnectionComponent(
                element,
                globalComponentIndex++,
                slideNumber - 1,
                element.zIndex,
                { debug }
              );
            }
            
              if (component) {
                const coords = component as any;
                if (coords.x > 50000 || coords.y > 50000 || coords.width > 50000 || coords.height > 50000) {
                  console.warn('Component coordinates may be in EMU, not pixels:', {
                    type: component.type,
                    x: coords.x,
                    y: coords.y,
                    width: coords.width,
                    height: coords.height
                  });
                }
                slideComponents.push(component);
                components.push(component);
                localComponentIndex++;
              }
          }
        } else {
          // Fallback to old method if ordered elements not available
          // Process text components
          for (const textComponent of slide.text) {
            const component = await this.parseUnifiedTextComponent(
              textComponent,
              globalComponentIndex++,
              slideNumber,
              localComponentIndex, // fallback zIndex based on order
              { debug, placeholders: combinedPlaceholders, masterStyles }
            );
            if (component) {
              slideComponents.push(component);
              components.push(component);
              localComponentIndex++;
            }
          }
          
          // Process shape components (non-text)
          for (const shapeComponent of slide.shapes) {
            const component = await this.parseUnifiedShapeComponent(
              shapeComponent,
              normalized.relationships,
              normalized.mediaFiles, 
              globalComponentIndex++,
              slideNumber - 1, // relationships index for media lookup
              localComponentIndex,
              { debug, placeholders: combinedPlaceholders }
            );
            if (component) {
              slideComponents.push(component);
              components.push(component);
              localComponentIndex++;
            }
          }
          
          // Process image components
          for (const imageComponent of slide.images) {
            const component = await this.parseUnifiedImageComponent(
              imageComponent,
              normalized.relationships,
              normalized.mediaFiles,
              globalComponentIndex++,
              slideNumber - 1,
              localComponentIndex,
              { debug, r2Storage, slideRelsFile }
            );
            if (component) {
              slideComponents.push(component);
              components.push(component);
              localComponentIndex++;
            }
          }


          // Process video components
          for (const videoComponent of slide.videos) {
            const component = await this.parseUnifiedVideoComponent(
              videoComponent,
              normalized.relationships,
              normalized.mediaFiles,
              globalComponentIndex++,
              slideNumber - 1,
              localComponentIndex,
              { debug, r2Storage, slideRelsFile }
            );
            if (component) {
              slideComponents.push(component);
              components.push(component);
              localComponentIndex++;
            }
          }
        }

        // Create slide object
        const slideObject: ParsedSlide = {
          slideIndex: slideNumber - 1, // For compatibility, keep 0-based index
          slideNumber: slideNumber,    // Position in presentation order
          title: slide.title,
          notes: slide.notes,
          layoutId,
          background: slideBackground,
          components: slideComponents,
          metadata: {
            name: slide.title || `Slide ${slideNumber}`,
            componentCount: slideComponents.length,
            format: normalized.format,
            slideFile: slide.slideFile || null,
            layoutFile: slide.layoutFile || null,
            masterFile: slide.masterFile || null,
            layoutElementCount: slide.layoutElementCount || 0,
            masterElementCount: slide.masterElementCount || 0,
            notesFile: slide.notesFile ?? null,
            fileSlideNumber: slide.fileSlideNumber ?? null,
            titleSource: slide.titleSource ?? null
          }
        };
        
        slides.push(slideObject);
        
        if (debug) {
          // console.log(`Slide ${slideNumber} complete: ${slideComponents.length} components`);
        }
      }
      
      if (debug) {
        // console.log('Unified parsing complete:', components.length, 'total components in', slides.length, 'slides');
      }
      
      // Validate slide dimensions are in pixel range before returning to client
      if (normalized.slideDimensions) {
        const { width, height } = normalized.slideDimensions;
        if (width > 50000 || height > 50000) {
          console.warn('Slide dimensions appear to be in EMU, not pixels:', { width, height });
        } else {
          // Slide dimensions validated as pixels
        }
      }

      // Clear theme colors after parsing
      BaseParser.clearThemeColors();

      // Fix connection points for all slides
      const fixedSlides = slides.map(slide => ({
        ...slide,
        components: this.fixConnectionPoints(slide.components)
      }));

      return {
        slides: fixedSlides,
        masters,
        layouts,
        totalComponents: components.length,
        format: normalized.format,
        slideOrderSource: normalized.slideOrderSource,
        slideDimensions: normalized.slideDimensions
      };

    } catch (error) {
      // Clear theme colors on error too
      BaseParser.clearThemeColors();
      console.error('❌ Error processing PowerPoint JSON:', error);
      throw error;
    }
  }

  /**
   * Parse a layout/master element to a PowerPoint component
   */
  private async parseElementToComponent(
    element: any,
    relationships: RelationshipGraph,
    mediaFiles: MediaFiles,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    options: { debug?: boolean; r2Storage?: R2BucketLike | null; slideRelsFile?: string | null } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, r2Storage = null, slideRelsFile = null } = options;

    if (debug) {
      // console.warn(`parseElementToComponent called with element.type: ${element.type}`);
      // console.warn(`Element data keys:`, Object.keys(element.data || {}));
    }
    
    // Convert LayoutElement to NormalizedElement format - strip namespace prefixes from data
    const strippedData = {};
    if (element.data && typeof element.data === 'object') {
      for (const [key, value] of Object.entries(element.data)) {
        // Strip namespace prefix (e.g., 'p:nvSpPr' -> 'nvSpPr')
        const strippedKey = key.includes(':') ? key.split(':')[1] : key;
        strippedData[strippedKey] = value;
      }
    }
    
    if (debug) {
      // console.warn(`Stripped data keys:`, Object.keys(strippedData));
    }
    
    const normalizedElement = {
      type: element.type,
      namespace: 'p' as const,
      element: element.type === 'image' ? 'pic' : 'sp',
      data: element.data, // Keep original data for compatibility
      zIndex: element.zIndex || zIndex,
      isLayoutElement: element.isLayoutElement,
      isMasterElement: element.isMasterElement,
      isBackgroundElement: element.isBackgroundElement,
      // Spread the stripped data properties at the top level
      ...strippedData
    };
    
    if (element.type === 'text' && isTextElement(normalizedElement)) {
      return await this.parseUnifiedTextComponent(
        normalizedElement,
        componentIndex,
        slideIndex,
        zIndex,
        { debug }
      );
    } else if (element.type === 'shape' && isShapeElement(normalizedElement)) {
      return await this.parseUnifiedShapeComponent(
        normalizedElement,
        relationships,
        mediaFiles,
        componentIndex,
        slideIndex,
        zIndex,
        { debug }
      );
    } else if (element.type === 'image' && isImageElement(normalizedElement)) {
      return await this.parseUnifiedImageComponent(
        normalizedElement,
        relationships,
        mediaFiles,
        componentIndex,
        slideIndex,
        zIndex,
        { debug, r2Storage, slideRelsFile }
      );
    } else if (element.type === 'table' && isTableElement(normalizedElement)) {
      return await this.parseUnifiedTableComponent(
        normalizedElement,
        relationships,
        mediaFiles,
        componentIndex,
        slideIndex,
        zIndex,
        { debug }
      );
    } else if (element.type === 'diagram' && isDiagramElement(normalizedElement)) {
      const diagramResult = await this.parseUnifiedDiagramComponent(
        normalizedElement,
        relationships,
        mediaFiles,
        componentIndex,
        slideIndex,
        zIndex,
        { debug, slideRelsFile }
      );
      // For parseElement, return the first component if it's an array
      if (Array.isArray(diagramResult)) {
        return diagramResult[0] || null;
      }
      return diagramResult;
    } else if (element.type === 'video' && isVideoElement(normalizedElement)) {
      return await this.parseUnifiedVideoComponent(
        normalizedElement,
        relationships,
        mediaFiles,
        componentIndex,
        slideIndex,
        zIndex,
        { debug, r2Storage, slideRelsFile }
      );
    } else if (element.type === 'connection' && isConnectionElement(normalizedElement)) {
      return await this.parseUnifiedConnectionComponent(
        normalizedElement,
        componentIndex,
        slideIndex,
        zIndex,
        { debug }
      );
    }
    
    if (debug) {
      // console.warn(`parseElementToComponent failed - element type: ${element.type}`);
      // console.warn(`Element data keys:`, Object.keys(element.data || {}));
      // console.warn(`normalizedElement.type:`, normalizedElement.type);
      // console.warn(`isShapeElement check:`, element.type === 'shape' ? isShapeElement(normalizedElement) : 'not shape');
      // console.warn(`isImageElement check:`, element.type === 'image' ? isImageElement(normalizedElement) : 'not image');
    }
    return null;
  }

  /**
   * Parse unified text component from normalized data
   */
  private async parseUnifiedTextComponent(
    textComponent: NormalizedTextElement,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    options: { debug?: boolean; placeholders?: PlaceholderMap; masterStyles?: { titleStyle?: XMLNode, bodyStyle?: XMLNode } } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, placeholders, masterStyles } = options;
    try {
      return await TextParser.parseFromNormalized(textComponent, componentIndex, slideIndex, zIndex, placeholders, masterStyles);
    } catch (error) {
      if (debug) console.warn(`Failed to parse text component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified shape component from normalized data
   */
  private async parseUnifiedShapeComponent(
    shapeComponent: NormalizedShapeElement,
    _relationships: RelationshipGraph,
    _mediaFiles: MediaFiles,
    componentIndex: number,
    relSlideIndex: number,
    zIndex: number,
    options: { debug?: boolean; placeholders?: PlaceholderMap; isMasterOrLayout?: boolean } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, placeholders, isMasterOrLayout = false } = options;
    try {
      return await ShapeParser.parseFromNormalized(shapeComponent, componentIndex, relSlideIndex, zIndex, placeholders, { isMasterOrLayout });
    } catch (error) {
      if (debug) console.warn(`Failed to parse shape component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified image component from normalized data
   */
  private async parseUnifiedImageComponent(
    imageComponent: NormalizedImageElement,
    relationships: RelationshipGraph,
    mediaFiles: MediaFiles,
    componentIndex: number,
    relSlideIndex: number,
    zIndex: number,
    options: { debug?: boolean; r2Storage?: R2BucketLike | null; slideRelsFile?: string | null } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, r2Storage = null, slideRelsFile = null } = options;
    try {
      return await ImageParser.parseFromNormalized(imageComponent, relationships, mediaFiles, componentIndex, relSlideIndex, zIndex, r2Storage, slideRelsFile);
    } catch (error) {
      if (debug) console.warn(`Failed to parse image component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified table component from normalized data
   */
  private async parseUnifiedTableComponent(
    tableComponent: NormalizedTableElement,
    _relationships: RelationshipGraph,
    _mediaFiles: MediaFiles,
    componentIndex: number,
    relSlideIndex: number,
    zIndex: number,
    options: { debug?: boolean } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false } = options;
    try {
      return await TableParser.parseFromNormalized(tableComponent, componentIndex, relSlideIndex, zIndex);
    } catch (error) {
      if (debug) console.warn(`Failed to parse table component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified diagram component from normalized data
   */
  private async parseUnifiedDiagramComponent(
    diagramComponent: NormalizedDiagramElement,
    _relationships: RelationshipGraph,
    _mediaFiles: MediaFiles,
    componentIndex: number,
    relSlideIndex: number,
    zIndex: number,
    options: { debug?: boolean; slideRelsFile?: string | null } = {}
  ): Promise<PowerPointComponent | PowerPointComponent[] | null> {
    const { debug = false, slideRelsFile = null } = options;
    try {
      return await DiagramParser.parseFromNormalized(diagramComponent, componentIndex, relSlideIndex, zIndex, this.normalizedJson, { returnExtractedComponents: true, slideRelsFile });
    } catch (error) {
      if (debug) console.warn(`Failed to parse diagram component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified video component from normalized data
   */
  private async parseUnifiedVideoComponent(
    videoComponent: NormalizedVideoElement, 
    _relationships: RelationshipGraph,
    _mediaFiles: MediaFiles,
    componentIndex: number, 
    relSlideIndex: number,
    zIndex: number,
    options: { debug?: boolean; r2Storage?: R2BucketLike | null; slideRelsFile?: string | null } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, r2Storage = null, slideRelsFile = null } = options;
    try {
      return await VideoParser.parseFromNormalized(videoComponent, _relationships, _mediaFiles, componentIndex, relSlideIndex, zIndex, r2Storage, slideRelsFile);
    } catch (error) {
      if (debug) console.warn(`Failed to parse video component:`, error);
      return null;
    }
  }
  
  /**
   * Parse unified connection component from normalized data
   */
  private async parseUnifiedConnectionComponent(
    connectionComponent: NormalizedConnectionElement,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    options: { debug?: boolean, shapeMap?: Map<string, { x: number, y: number, width: number, height: number }> } = {}
  ): Promise<PowerPointComponent | null> {
    const { debug = false, shapeMap } = options;
    try {
      return await ConnectorParser.parseFromNormalized(connectionComponent, componentIndex, slideIndex, zIndex, shapeMap);
    } catch (error) {
      if (debug) console.warn(`Failed to parse connection component:`, error);
      return null;
    }
  }

  /**
   * Post-process connection components to fix connection points
   * This corrects startPoint and endPoint based on actual connected shape positions
   * @param components - Array of all parsed components
   * @returns Updated components with corrected connection points
   */
  private fixConnectionPoints(components: PowerPointComponent[]): PowerPointComponent[] {
    // Build shape lookup map
    const shapeMap = new Map<string, { x: number, y: number, width: number, height: number }>();
    
    for (const component of components) {
      if (component.type !== 'connection' && component.id) {
        // Extract shape ID from component ID (format: "Name;ID;namespace")
        const idParts = component.id.split(';');
        if (idParts.length >= 2) {
          const shapeId = idParts[1];
          shapeMap.set(shapeId, {
            x: component.x,
            y: component.y,
            width: component.width,
            height: component.height
          });
        }
      }
    }

    // Fix connection points
    return components.map(component => {
      if (component.type === 'connection') {
        const connection = component as ConnectionComponent;
        
        if (connection.startShapeId && connection.endShapeId) {
          const startShape = shapeMap.get(connection.startShapeId);
          const endShape = shapeMap.get(connection.endShapeId);
          
          if (startShape && endShape) {
            // Use ConnectorParser to calculate proper connection points based on actual indices
            // Extract connection indices from metadata or default to right->left connection
            const startConnectionIndex = connection.metadata?.startConnectionIndex ?? 3; // Default to right center
            const endConnectionIndex = connection.metadata?.endConnectionIndex ?? 1; // Default to left center
            
            const startPoint = ConnectorParser.calculateConnectionPoint(startShape, startConnectionIndex);
            const endPoint = ConnectorParser.calculateConnectionPoint(endShape, endConnectionIndex);
            
            return {
              ...connection,
              startPoint,
              endPoint
            };
          }
        }
      }
      return component;
    });
  }

  /**
   * Extract text styles (txStyles) from a slide master
   * @param json - PPTXJson object containing all parsed files
   * @param masterFile - Path to the master file (e.g., 'ppt/slideMasters/slideMaster1.xml')
   * @returns Master text styles object with titleStyle and bodyStyle
   */
  private extractMasterTextStyles(json: any, masterFile: string): { titleStyle?: XMLNode, bodyStyle?: XMLNode } {
    if (!json || !json[masterFile]) {
      return {};
    }

    const masterData = json[masterFile];
    return this.extractStylesFromMasterData(masterData);
  }

  /**
   * Helper method to extract styles from master data
   */
  private extractStylesFromMasterData(masterData: any): { titleStyle?: XMLNode, bodyStyle?: XMLNode } {
    // Navigate to txStyles: sldMaster > txStyles > titleStyle/bodyStyle
    const sldMaster = BaseParser.getNode(masterData, 'sldMaster');
    if (!sldMaster) {
      // Try alternative namespace-aware paths
      const altSldMaster = BaseParser.getNode(masterData, 'p:sldMaster') || masterData['p:sldMaster'];
      if (!altSldMaster) {
        return {};
      }
      const altTxStyles = BaseParser.getNode(altSldMaster, 'txStyles') || BaseParser.getNode(altSldMaster, 'p:txStyles');
      if (!altTxStyles) {
        return {};
      }
      
      const titleStyle = BaseParser.getNode(altTxStyles, 'titleStyle') || BaseParser.getNode(altTxStyles, 'p:titleStyle');
      const bodyStyle = BaseParser.getNode(altTxStyles, 'bodyStyle') || BaseParser.getNode(altTxStyles, 'p:bodyStyle');
      
      return {
        titleStyle: titleStyle || undefined,
        bodyStyle: bodyStyle || undefined
      };
    }

    const txStyles = BaseParser.getNode(sldMaster, 'txStyles');
    if (!txStyles) {
      return {};
    }

    const titleStyle = BaseParser.getNode(txStyles, 'titleStyle');
    const bodyStyle = BaseParser.getNode(txStyles, 'bodyStyle');

    return {
      titleStyle: titleStyle || undefined,
      bodyStyle: bodyStyle || undefined
    };
  }
}