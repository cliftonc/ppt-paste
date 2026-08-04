/**
 * Diagram component parser for PowerPoint diagrams and SmartArt
 */

import { BaseParser, TransformInfo } from "./BaseParser.js";
import { ShapeParser } from "./ShapeParser.js";
import { emuToPixels } from "../utils/constants.js";
import {
  XMLNode,
  DiagramComponent,
  SmartArtDataPoint,
  SmartArtConnection,
  SmartArtShape,
  SmartArtLayout,
  PowerPointComponent,
  TextComponent,
  ShapeComponent,
} from "../types/index.js";
import { NormalizedDiagramElement } from "../types/normalized.js";

export class DiagramParser extends BaseParser {
  /**
   * Parse diagram component from normalized data
   * @param diagramComponent - Normalized diagram component
   * @param componentIndex - Component index
   * @param slideIndex - Slide index
   * @param zIndex - Z-index for layering
   * @param parsedFiles - All parsed PPTX files for accessing diagram data
   * @param options - Parsing options
   * @returns Parsed diagram component or array of extracted components
   */
  static async parseFromNormalized(
    diagramComponent: NormalizedDiagramElement,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    parsedFiles?: Record<string, XMLNode | Uint8Array | string>,
    options: { returnExtractedComponents?: boolean; slideRelsFile?: string | null } = {}
  ): Promise<DiagramComponent | PowerPointComponent[] | null> {
    const { spPr, nvGraphicFramePr, graphicData } = diagramComponent;

    if (!spPr) {
      return null;
    }

    // Extract positioning from spPr (which is actually xfrm for graphicFrame)
    const transform = DiagramParser.parseTransform(spPr);

    // Skip if diagram has no dimensions
    if (transform.width === 0 && transform.height === 0) return null;

    // Extract component info from nvGraphicFramePr
    const originalName = BaseParser.getString(nvGraphicFramePr, "cNvPr.$name", "");
    const componentName = BaseParser.generateComponentId('diagram', componentIndex, originalName);

    // Try to determine diagram type from graphicData
    let diagramType = "unknown";
    let title = componentName;
    let smartArtData = undefined;
    let extractedComponents: PowerPointComponent[] = [];
    
    if (graphicData) {
      // Look for diagram relationships that might give us clues about the type
      const relIds = graphicData.relIds || graphicData;
      if (relIds) {
        diagramType = "smartart"; // Most diagrams in PowerPoint are SmartArt
        
        // Extract SmartArt data if we have access to parsed files
        if (parsedFiles) {
          try {
            smartArtData = await DiagramParser.extractSmartArtData(
              relIds,
              parsedFiles,
              slideIndex,
              options.slideRelsFile ?? null
            );
            if (smartArtData) {
              // Generate individual components from SmartArt data
              extractedComponents = DiagramParser.generateComponentsFromSmartArt(
                smartArtData,
                slideIndex,
                zIndex,
                transform.x,
                transform.y
              );
            }
          } catch (error) {
            console.warn("Failed to extract SmartArt data:", error);
          }
        }
      }
    }

    const result: DiagramComponent = {
      id: componentName,
      type: "diagram",
      content: title,
      x: transform.x,
      y: transform.y,
      width: transform.width,
      height: transform.height,
      rotation: transform.rotation || 0,
      slideIndex,
      zIndex,
      diagramType,
      title,
      style: {
        opacity: 1,
        rotation: transform.rotation || 0,
      },
      metadata: {
        namespace: diagramComponent.namespace,
        originalFormat: "normalized",
        diagramType,
        hasGraphicData: !!graphicData,
      },
    };
    
    // Add SmartArt data if extracted
    if (smartArtData) {
      result.smartArtData = smartArtData;
    }
    
    // Add extracted components if any
    if (extractedComponents.length > 0) {
      result.extractedComponents = extractedComponents;
      
      // If option is set, return only the extracted components instead of the diagram wrapper
      if (options.returnExtractedComponents) {
        return extractedComponents;
      }
    }

    return result;
  }

  /**
   * Parse transform information from xfrm element
   * @param xfrm - Transform element
   * @returns Transform data with position and size
   */
  static parseTransform(xfrm: XMLNode): TransformInfo {
    if (!xfrm) {
      return { x: 0, y: 0, width: 0, height: 0, rotation: 0 };
    }

    // Extract offset (position)
    const off = BaseParser.getNode(xfrm, "off");
    const x = off ? emuToPixels(BaseParser.getNumber(off, "$x", 0)) : 0;
    const y = off ? emuToPixels(BaseParser.getNumber(off, "$y", 0)) : 0;

    // Extract extent (size)
    const ext = BaseParser.getNode(xfrm, "ext");
    const width = ext ? emuToPixels(BaseParser.getNumber(ext, "$cx", 0)) : 0;
    const height = ext ? emuToPixels(BaseParser.getNumber(ext, "$cy", 0)) : 0;

    // Extract rotation if present
    const rotation = BaseParser.getNumber(xfrm, "$rot", 0);
    const rotationDegrees = rotation ? rotation / 60000 : 0; // PowerPoint uses 60000 units per degree

    return {
      x,
      y,
      width,
      height,
      rotation: rotationDegrees,
    };
  }

  /**
   * Extract SmartArt data from diagram relationship IDs
   */
  static async extractSmartArtData(
    relIds: any,
    parsedFiles: Record<string, XMLNode | Uint8Array | string>,
    currentSlideIndex?: number,
    slideRelsFile?: string | null
  ): Promise<{
    dataPoints: SmartArtDataPoint[];
    connections: SmartArtConnection[];
    shapes: SmartArtShape[];
    layout: SmartArtLayout;
  } | null> {
    try {
      // Get diagram-specific files from the relationship data
      // The relIds should contain relationship information that maps to specific diagram files
      
      let specificDataFiles: string[] = [];
      let specificDrawingFiles: string[] = [];
      
      // First, try to get the specific files from relationship IDs
      // If we can't determine specific files, fall back to all diagram files
      if (relIds && typeof relIds === 'object') {
        // Look through the graphicData for relationship references
        const relationsFound = this.extractRelationshipTargets(
          relIds,
          parsedFiles,
          currentSlideIndex,
          slideRelsFile
        );
        
        if (relationsFound.dataFiles.length > 0 || relationsFound.drawingFiles.length > 0) {
          specificDataFiles = relationsFound.dataFiles;
          specificDrawingFiles = relationsFound.drawingFiles;
        }
      }
      
      // Fallback: if we couldn't determine specific files, get all diagram files
      if (specificDataFiles.length === 0 && specificDrawingFiles.length === 0) {
        console.warn("DiagramParser: Could not determine specific diagram files from relIds, falling back to all diagram files");
        specificDataFiles = Object.keys(parsedFiles).filter(path => 
          path.includes('/diagrams/data') && path.endsWith('.xml')
        );
        specificDrawingFiles = Object.keys(parsedFiles).filter(path =>
          path.includes('/diagrams/drawing') && path.endsWith('.xml')
        );
      }
      
      if (specificDataFiles.length === 0 && specificDrawingFiles.length === 0) {
        return null;
      }
      
      let dataPoints: SmartArtDataPoint[] = [];
      let connections: SmartArtConnection[] = [];
      let layout: SmartArtLayout = { layoutType: 'unknown' };
      let shapes: SmartArtShape[] = [];
      
      // Parse data model files
      for (const dataFile of specificDataFiles) {
        const dataXml = parsedFiles[dataFile] as XMLNode;
        if (!dataXml) continue;
        
        const dataModel = dataXml['dataModel'];
        if (dataModel) {
          const extractedData = DiagramParser.parseSmartArtDataModel(dataModel);
          dataPoints.push(...extractedData.dataPoints);
          connections.push(...extractedData.connections);
          if (extractedData.layout) {
            layout = { ...layout, ...extractedData.layout };
          }
        }
      }
      
      // Parse drawing files
      for (const drawingFile of specificDrawingFiles) {
        const drawingXml = parsedFiles[drawingFile] as XMLNode;
        if (!drawingXml) continue;
        
        const drawing = drawingXml['drawing'];
        if (drawing) {
          const extractedShapes = DiagramParser.parseSmartArtDrawing(drawing);
          shapes.push(...extractedShapes);
        }
      }
      
      return { dataPoints, connections, shapes, layout };
      
    } catch (error) {
      console.warn("Error extracting SmartArt data:", error);
      return null;
    }
  }

  /**
   * Extract relationship targets from diagram graphicData
   */
  static extractRelationshipTargets(
    relIds: any,
    parsedFiles: Record<string, XMLNode | Uint8Array | string>,
    currentSlideIndex?: number,
    slideRelsFile?: string | null
  ): {
    dataFiles: string[];
    drawingFiles: string[];
  } {
    const dataFiles: string[] = [];
    const drawingFiles: string[] = [];
    
    // Extract relationship IDs from the relIds object
    const relationshipIds: string[] = [];
    if (relIds.dm) relationshipIds.push(relIds.dm); // Data model
    if (relIds.lo) relationshipIds.push(relIds.lo); // Layout
    if (relIds.qs) relationshipIds.push(relIds.qs); // QuickStyle  
    if (relIds.cs) relationshipIds.push(relIds.cs); // Color scheme
    if (relIds.drawing) relationshipIds.push(relIds.drawing); // Drawing
    
    // Look for additional relationship references in the object
    // Sometimes the drawing reference might be in a different field
    for (const [key, value] of Object.entries(relIds)) {
      if (typeof value === 'string' && value.startsWith('rId') && !relationshipIds.includes(value)) {
        relationshipIds.push(value);
      }
    }
    
    // Find the correct slide relationship file. The slide's own rels part is passed in
    // wherever it is known; deriving it from the slide index assumes presentation order
    // matches the slide's FILE number, which a reordered deck breaks.
    const slideRelsPath =
      slideRelsFile || `ppt/slides/_rels/slide${(currentSlideIndex || 0) + 1}.xml.rels`;
    
    const slideRels = parsedFiles[slideRelsPath] as XMLNode;
    if (slideRels && slideRels['Relationships']) {
      const relationships = slideRels['Relationships']['Relationship'];
      const relationshipsArray = Array.isArray(relationships) ? relationships : [relationships];
      
      // Map relationship IDs to target files
      for (const relationshipId of relationshipIds) {
        const relationship = relationshipsArray.find((rel: any) => 
          rel && rel['$Id'] === relationshipId
        );
        
        if (relationship && relationship['$Target']) {
          const target = relationship['$Target'];
          
          // Convert relative path to full path
          // Target is relative to the slide, e.g., "../diagrams/data1.xml"
          let fullPath = target;
          if (target.startsWith('../')) {
            fullPath = `ppt/${target.substring(3)}`;
          }
          
          if (fullPath.includes('/diagrams/data') && fullPath.endsWith('.xml')) {
            dataFiles.push(fullPath);
          } else if (fullPath.includes('/diagrams/drawing') && fullPath.endsWith('.xml')) {
            drawingFiles.push(fullPath);
          }
        }
      }
      
      // Also check for diagram drawing files in all relationships on this slide
      // as the drawing relationship might not be in the explicit relIds
      for (const relationship of relationshipsArray) {
        if (relationship && relationship['$Target']) {
          const target = relationship['$Target'];
          let fullPath = target;
          if (target.startsWith('../')) {
            fullPath = `ppt/${target.substring(3)}`;
          }
          
          // If this is a drawing file and we haven't found it yet, add it
          if (fullPath.includes('/diagrams/drawing') && fullPath.endsWith('.xml') && !drawingFiles.includes(fullPath)) {
            drawingFiles.push(fullPath);
          }
        }
      }
    }
    return { dataFiles, drawingFiles };
  }

  /**
   * Parse SmartArt data model from dgm:dataModel
   */
  static parseSmartArtDataModel(dataModel: XMLNode): {
    dataPoints: SmartArtDataPoint[];
    connections: SmartArtConnection[];
    layout: Partial<SmartArtLayout>;
  } {
    const dataPoints: SmartArtDataPoint[] = [];
    const connections: SmartArtConnection[] = [];
    const layout: Partial<SmartArtLayout> = {};
    
    // Parse point list
    const ptLst = dataModel['ptLst'];
    if (ptLst && ptLst['pt']) {
      const points = Array.isArray(ptLst['pt']) ? ptLst['pt'] : [ptLst['pt']];
      
      for (const point of points) {
        const modelId = BaseParser.getString(point, '$modelId', '');
        const type = BaseParser.getString(point, '$type', '');
        
        // Extract text content
        let content = '';
        const textNode = point['t'];
        if (textNode) {
          const paragraph = textNode['p'];
          if (paragraph && paragraph['r']) {
            const runs = Array.isArray(paragraph['r']) ? paragraph['r'] : [paragraph['r']];
            content = runs.map((run: XMLNode) => 
              BaseParser.getString(run, 't', '')
            ).join('');
          }
        }
        
        // Extract layout info from first point
        if (type === 'doc') {
          const prSet = point['prSet'];
          if (prSet) {
            layout.layoutType = BaseParser.getString(prSet, '$loTypeId', '');
            layout.layoutCategory = BaseParser.getString(prSet, '$loCatId', '');
            layout.colorScheme = BaseParser.getString(prSet, '$csTypeId', '');
            layout.quickStyle = BaseParser.getString(prSet, '$qsTypeId', '');
          }
        }
        
        dataPoints.push({
          modelId,
          type,
          content: content.trim(),
          children: [],
          connections: [],
        });
      }
    }
    
    // Parse connection list
    const cxnLst = dataModel['cxnLst'];
    if (cxnLst && cxnLst['cxn']) {
      const cxns = Array.isArray(cxnLst['cxn']) ? cxnLst['cxn'] : [cxnLst['cxn']];
      
      for (const cxn of cxns) {
        const connection: SmartArtConnection = {
          connectionId: BaseParser.getString(cxn, '$modelId', ''),
          type: BaseParser.getString(cxn, '$type', 'default'),
          sourceId: BaseParser.getString(cxn, '$srcId', ''),
          destinationId: BaseParser.getString(cxn, '$destId', ''),
          sourceOrder: BaseParser.getNumber(cxn, '$srcOrd', 0),
          destinationOrder: BaseParser.getNumber(cxn, '$destOrd', 0),
        };
        
        connections.push(connection);
      }
    }
    
    return { dataPoints, connections, layout };
  }

  /**
   * Parse SmartArt drawing to extract visual shapes
   */
  static parseSmartArtDrawing(drawing: XMLNode): SmartArtShape[] {
    const shapes: SmartArtShape[] = [];
    
    const spTree = drawing['spTree'];
    if (!spTree) return shapes;
    
    // Find all shape elements
    const shapeElements = spTree['sp'];
    if (!shapeElements) return shapes;
    
    const shapeArray = Array.isArray(shapeElements) ? shapeElements : [shapeElements];
    
    for (const shape of shapeArray) {
      const modelId = BaseParser.getString(shape, '$modelId', '');
      
      // Extract transform
      const spPr = shape['spPr'];
      if (!spPr) continue;
      
      const transform = DiagramParser.parseTransform(spPr['xfrm']);
      
      // Extract geometry
      const prstGeom = spPr['prstGeom'];
      const shapeType = BaseParser.getString(prstGeom, '$prst', 'rectangle');
      
      // Extract fill
      const fillInfo = ShapeParser.parseFill(spPr, null);
      
      // Extract border
      const borderInfo = ShapeParser.parseBorder(spPr, null);
      
      // Extract text content
      let textContent = '';
      let textStyle = {};
      
      const txBody = shape['txBody'];
      if (txBody) {
        const paragraph = txBody['p'];
        if (paragraph && paragraph['r']) {
          const runs = Array.isArray(paragraph['r']) ? paragraph['r'] : [paragraph['r']];
          textContent = runs.map((run: XMLNode) => 
            BaseParser.getString(run, 't', '')
          ).join('');
          
          // Extract text style from first run
          if (runs[0] && runs[0]['rPr']) {
            const rPr = runs[0]['rPr'];
            textStyle = {
              fontSize: BaseParser.getNumber(rPr, '$sz', 0) / 100, // Convert from points * 100
              fontFamily: BaseParser.getString(rPr, '$typeface', undefined),
              color: BaseParser.parseColor(BaseParser.getNode(rPr, 'solidFill')),
            };
          }
        }
      }
      
      shapes.push({
        modelId,
        shapeType,
        x: transform.x,
        y: transform.y,
        width: transform.width,
        height: transform.height,
        rotation: transform.rotation,
        fill: fillInfo,
        border: borderInfo,
        effects: [],
        textContent: textContent.trim(),
        textStyle,
      });
    }
    
    return shapes;
  }

  /**
   * Generate individual components from SmartArt data
   */
  static generateComponentsFromSmartArt(
    smartArtData: {
      dataPoints: SmartArtDataPoint[];
      connections: SmartArtConnection[];
      shapes: SmartArtShape[];
      layout: SmartArtLayout;
    },
    slideIndex: number,
    baseZIndex: number,
    offsetX: number,
    offsetY: number
  ): PowerPointComponent[] {
    const components: PowerPointComponent[] = [];
    let currentZIndex = baseZIndex;
    
    // Match data points with shapes and create components
    // The order in the shapes array represents the intended layering from the XML
    for (let i = 0; i < smartArtData.shapes.length; i++) {
      const shape = smartArtData.shapes[i];
      
      // Find corresponding data point
      const dataPoint = smartArtData.dataPoints.find(dp => dp.modelId === shape.modelId);
      
      // Determine if this should be a text or shape component
      const hasText = shape.textContent && shape.textContent.length > 0;
      
      if (hasText) {
        // Create text component matching TextParser structure
        const textComponent: TextComponent = {
          id: `smartart-text-${shape.modelId}`,
          type: 'text',
          content: shape.textContent || '',
          x: shape.x + offsetX,
          y: shape.y + offsetY,
          width: shape.width,
          height: shape.height,
          rotation: shape.rotation || 0,
          slideIndex,
          zIndex: ++currentZIndex,
          style: {
            fontSize: shape.textStyle?.fontSize || 18,
            fontFamily: shape.textStyle?.fontFamily || 'Arial',
            fontWeight: 'normal',
            fontStyle: 'normal',
            textDecoration: 'none',
            color: shape.textStyle?.color || '#000000',
            backgroundColor: 'transparent',
            textAlign: 'center',
            opacity: 1,
            rotation: shape.rotation || 0
          },
          richText: {
            type: 'doc',
            content: [{
              type: 'paragraph',
              content: [{
                type: 'text',
                text: shape.textContent || '',
                marks: shape.textStyle?.fontSize ? [{
                  type: 'textStyle',
                  attrs: { fontSize: `${shape.textStyle.fontSize}pt` }
                }] : []
              }]
            }]
          },
          metadata: {
            namespace: 'smartart',
            isTextBox: false,
            originalFormat: 'smartart',
            paragraphCount: 1,
            hasMultipleRuns: false,
            smartArtModelId: shape.modelId,
            dataPointContent: dataPoint?.content,
            layoutType: smartArtData.layout.layoutType,
          },
        };
        
        // Add background shape if shape has fill/border
        if (shape.fill || shape.border) {
          textComponent.backgroundShape = {
            type: shape.shapeType === 'roundRect' ? 'roundRect' : 'rectangle',
            fill: shape.fill,
            border: shape.border,
          };
        }
        
        components.push(textComponent);
      } else {
        // Create shape component matching ShapeParser structure
        const geometry = {
          type: shape.shapeType === 'roundRect' ? 'rounded rectangle' : 'rectangle',
          preset: shape.shapeType,
          isCustom: false,
        };
        
        const shapeComponent: ShapeComponent = {
          id: `smartart-shape-${shape.modelId}`,
          type: 'shape',
          content: `${geometry.type} shape`,
          x: shape.x + offsetX,
          y: shape.y + offsetY,
          width: shape.width,
          height: shape.height,
          rotation: shape.rotation || 0,
          slideIndex,
          zIndex: ++currentZIndex,
          shapeType: geometry.type,
          geometry: geometry,
          style: {
            fillColor: shape.fill?.color || 'transparent',
            borderColor: shape.border?.color || 'transparent',
            borderWidth: shape.border?.width || 0,
            borderStyle: shape.border?.style || 'none',
            fillOpacity: shape.fill?.opacity || 1,
            rotation: shape.rotation || 0,
            effects: shape.effects || [],
          },
          metadata: {
            namespace: 'smartart',
            geometry: geometry,
            originalFormat: 'smartart',
            shapeType: geometry.type,
            hasEffects: (shape.effects || []).length > 0,
            hasFill: !!(shape.fill?.color && shape.fill.color !== 'transparent'),
            hasBorder: !!(shape.border?.type && shape.border.type !== 'none'),
            smartArtModelId: shape.modelId,
            dataPointContent: dataPoint?.content,
            layoutType: smartArtData.layout.layoutType,
          },
        };
        
        components.push(shapeComponent);
      }
    }
    
    return components;
  }
}