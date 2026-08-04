/**
 * Image component parser for PowerPoint shapes containing images
 */

import { BaseParser, isBufferLike, bufferFrom } from "./BaseParser.js";
import {
  XMLNode,
  ImageComponent,
  NormalizedImageComponent,
  ImageInfo,
  ImageDimensions,
  ImageEffectsInfo,
  MediaFileInfo,
  ImageCroppingInfo,
} from "../types/index.js";

export class ImageParser extends BaseParser {
  /**
   * Parse image component from normalized data (works for both PPTX and clipboard)
   * @param imageComponent - Normalized image component
   * @param relationships - Relationship data
   * @param mediaFiles - Media files
   * @param componentIndex - Component index
   * @param slideIndex - Slide index (position in presentation order; also the component's own slideIndex)
   * @param r2Storage - Optional R2 storage for image hosting
   * @param slideRelsFile - The slide's own .rels part. Pass it whenever you have it:
   *   deriving the path from `slideIndex` assumes presentation order matches the slide's
   *   FILE number, which stops being true the moment a deck is reordered.
   * @returns Parsed image component
   */
  static async parseFromNormalized(
    imageComponent: NormalizedImageComponent,
    relationships: Record<string, any>,
    mediaFiles: Record<string, Uint8Array>,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    r2Storage: any = null,
    slideRelsFile: string | null = null,
  ): Promise<ImageComponent | null> {
    const { data, spPr, nvPicPr, blipFill, namespace } = imageComponent;

    if (!spPr || !blipFill) {
      throw new Error(
        "No spPr or blipFill found in normalized image component",
      );
    }

    // Extract positioning from spPr (namespaces already stripped)
    const xfrm = BaseParser.getNode(spPr, "xfrm");
    const transform = ImageParser.parseTransform(xfrm);

    // Extract component info from nvPicPr
    const originalName = BaseParser.getString(nvPicPr, "cNvPr.$name", "");
    const componentName = BaseParser.generateComponentId('image', componentIndex, originalName);
    const description = BaseParser.getString(nvPicPr, "cNvPr.$descr", "");

    // Extract image reference from blipFill (namespaces already stripped)
    const relationshipId = BaseParser.getString(blipFill, "blip.embed", "");

    // Find the actual image file using relationships
    let imageDataUrl: string | null = null;
    let imageFormat = "unknown";
    let imageSize = 0;

    if (relationshipId && relationships && mediaFiles) {
      // Use the more complete getImageInfo method with slide context
      const imageInfo = await ImageParser.getImageInfo(
        relationshipId,
        relationships,
        mediaFiles,
        slideIndex,
        r2Storage,
        slideRelsFile,
      );
      if (imageInfo.url) {
        imageDataUrl = imageInfo.url;
        imageFormat = imageInfo.type;
        imageSize = imageInfo.size;
      }
    }

    // Skip WMF/EMF files - these are often decorative background elements
    if (imageFormat === 'wmf' || imageFormat === 'emf') {
      return null;
    }

    // Parse image effects using existing method
    const effects = ImageParser.parseImageEffects(blipFill);

    return {
      id: componentName,
      type: "image",
      content: description || componentName,
      x: transform.x,
      y: transform.y,
      width: transform.width,
      height: transform.height,
      slideIndex,
      style: {
        rotation: transform.rotation || 0,
        fillOpacity: effects.opacity,
        ...effects,
      },
      src: imageDataUrl || "",
      alt: description || componentName,
      zIndex,
      metadata: {
        namespace,
        name: componentName,
        description,
        relationshipId,
        imageUrl: imageDataUrl,
        imageType: imageFormat,
        imageSize: imageSize,
        originalFormat: "normalized",
        hasEffects: effects.effectsList.length > 0,
      },
    };
  }

  /**
   * Find media file from relationships and mediaFiles
   */
  static findMediaFile(
    relationshipId: string,
    relationships: Record<string, any>,
    mediaFiles: Record<string, Uint8Array>,
  ): MediaFileInfo | null {
    if (!relationshipId || !relationships || !mediaFiles) {
      return null;
    }

    const relFile = Object.keys(relationships).find(
      (key) => key.includes("_rels") && key.includes("drawing"),
    );

    if (!relFile || !relationships[relFile]) {
      return null;
    }

    const relationshipArray = BaseParser.getArray(
      relationships[relFile],
      "Relationships.Relationship",
      [],
    );

    const relationshipData = relationshipArray.find(
      (rel: any) =>
        BaseParser.asString(rel?.$Id, "") ===
        BaseParser.asString(relationshipId, ""),
    );

    if (!relationshipData) {
      return null;
    }

    let mediaPath = BaseParser.asString(relationshipData.$Target, "");
    if (mediaPath.startsWith("../")) {
      mediaPath = `clipboard/${mediaPath.slice(3)}`;
    }

    const mediaFile = mediaFiles[mediaPath];
    if (mediaFile) {
      return {
        data: mediaFile,
        type: this.getImageTypeFromPath(mediaPath),
        size: (mediaFile as any).byteLength || (mediaFile as any).length || 0,
      };
    }

    return null;
  }

  /**
   * Get image type from file path
   */
  static getImageTypeFromPath(path: string): string {
    const ext = path.toLowerCase().split(".").pop();
    switch (ext) {
      case "png":
        return "image/png";
      case "jpg":
      case "jpeg":
        return "image/jpeg";
      case "gif":
        return "image/gif";
      case "svg":
        return "image/svg+xml";
      default:
        return "image/unknown";
    }
  }

  /**
   * Get file extension from MIME type
   */
  static getExtensionFromMimeType(mimeType: string): string {
    const extensions: Record<string, string> = {
      "image/jpeg": "jpg",
      "image/png": "png",
      "image/gif": "gif",
      "image/svg+xml": "svg",
      "image/bmp": "bmp",
      "image/webp": "webp",
      "image/tiff": "tiff",
    };
    return extensions[mimeType] || "png";
  }

  /**
   * Get image information from relationship and media data
   * @param rId - Relationship ID
   * @param relationships - Relationship data
   * @param mediaFiles - Media files data
   * @param slideIndex - Slide index for slide-scoped search
   * @param r2Storage - Optional R2 storage
   * @param slideRelsFile - The slide's own .rels part; preferred over deriving one from slideIndex
   * @returns image information
   */
  static async getImageInfo(
    rId: string,
    relationships: Record<string, any>,
    mediaFiles: Record<string, Uint8Array>,
    slideIndex: number | null = null,
    r2Storage: any = null,
    slideRelsFile: string | null = null,
  ): Promise<ImageInfo> {
    if (!rId) {
      return {
        url: null,
        type: "unknown",
        size: 0,
        dimensions: null,
      };
    }

    // Detect if this is clipboard format vs PPTX format by checking relationship paths
    const isClipboardFormat = Object.keys(relationships || {}).some((key) =>
      key.includes("clipboard/"),
    );

    // For PPTX files (not clipboard), search current slide first, then slide layouts/masters
    if ((slideRelsFile || slideIndex !== null) && !isClipboardFormat) {
      // Prefer the slide's actual rels part. `slideIndex + 1` is the slide's presentation
      // position, and the rels graph is keyed on the FILE number — the same thing only
      // until a deck is reordered.
      const currentSlideRelFile =
        slideRelsFile || `ppt/slides/_rels/slide${(slideIndex as number) + 1}.xml.rels`;

      // First try the current slide's relationship file
      if (relationships[currentSlideRelFile]) {
        const relsData =
          (BaseParser.safeGet(
            relationships[currentSlideRelFile],
            "Relationships.Relationship",
            [],
          ) as any) || [];
        const rels = Array.isArray(relsData) ? relsData : [relsData];
        const rel = rels.find(
          (r: any) =>
            BaseParser.asString(r?.$Id, "") === BaseParser.asString(rId, ""),
        );

        if (rel) {
          const target = BaseParser.asString(rel?.$Target, "");
          let mediaPath: string = "";

          if (target.startsWith("../")) {
            // PPTX format: ../media/image1.png -> ppt/media/image1.png
            mediaPath = `ppt/${target.slice(3)}`;
          } else {
            mediaPath = target.startsWith("media/") ? `ppt/${target}` : target;
          }

          // Look for the media file
          const mediaFile = mediaFiles[mediaPath];
          if (mediaFile) {
            return {
              url: await ImageParser.createImageUrl(
                mediaFile,
                target,
                r2Storage,
              ),
              type: ImageParser.getImageType(target),
              size: (mediaFile as any).length || 0,
              dimensions: ImageParser.getImageDimensions(mediaFile),
            };
          }
        }
      }

      // If not found in slide's relationships, also check slide masters and layouts
      // This is important for background images that come from masters
      const masterAndLayoutRelFiles = Object.keys(relationships).filter(
        (key) =>
          key.includes("_rels") &&
          (key.includes("slideMaster") || key.includes("slideLayout")),
      );

      for (const relFile of masterAndLayoutRelFiles) {
        if (relationships[relFile]) {
          const relsData =
            (BaseParser.safeGet(
              relationships[relFile],
              "Relationships.Relationship",
              [],
            ) as any) || [];
          const rels = Array.isArray(relsData) ? relsData : [relsData];
          const rel = rels.find(
            (r: any) =>
              BaseParser.asString(r?.$Id, "") === BaseParser.asString(rId, ""),
          );

          if (rel) {
            const target = BaseParser.asString(rel?.$Target, "");
            let mediaPath: string = "";

            if (target.startsWith("../")) {
              // PPTX format: ../media/image1.png -> ppt/media/image1.png
              mediaPath = `ppt/${target.slice(3)}`;
            } else {
              mediaPath = target.startsWith("media/")
                ? `ppt/${target}`
                : target;
            }

            // Look for the media file
            const mediaFile = mediaFiles[mediaPath];
            if (mediaFile) {
              return {
                url: await ImageParser.createImageUrl(
                  mediaFile,
                  target,
                  r2Storage,
                ),
                type: ImageParser.getImageType(target),
                size: (mediaFile as any).length || 0,
                dimensions: ImageParser.getImageDimensions(mediaFile),
              };
            }
          }
        }
      }

      // For PPTX with slideIndex, if not found anywhere, return null
      return {
        url: null,
        type: "unknown",
        size: 0,
        dimensions: null,
      };
    }

    // Only for clipboard format (when slideIndex is null) - search through all relationship files
    const relFiles = Object.keys(relationships).filter(
      (key) =>
        key.includes("_rels") &&
        (key.includes("slide") || key.includes("drawing")),
    );

    for (const relFile of relFiles) {
      if (relationships[relFile]) {
        const relsData =
          (BaseParser.safeGet(
            relationships[relFile],
            "Relationships.Relationship",
            [],
          ) as any) || [];
        const rels = Array.isArray(relsData) ? relsData : [relsData];
        const rel = rels.find(
          (r: any) =>
            BaseParser.asString(r?.$Id, "") === BaseParser.asString(rId, ""),
        );

        if (rel) {
          const target = BaseParser.asString(rel?.$Target, "");
          let mediaPath: string = "";

          if (target.startsWith("../")) {
            // Check if we're dealing with clipboard or PPTX format based on relationship file path
            if (relFile.includes("clipboard")) {
              mediaPath = `clipboard/${target.slice(3)}`;
            } else {
              // PPTX format - resolve relative path from slide to media folder
              mediaPath = `ppt/${target.slice(3)}`;
            }
          } else {
            mediaPath = target.startsWith("media/") ? `ppt/${target}` : target;
          }

          // Look for the media file
          const mediaFile = mediaFiles[mediaPath];
          if (mediaFile) {
            return {
              url: await ImageParser.createImageUrl(
                mediaFile,
                target,
                r2Storage,
              ),
              type: ImageParser.getImageType(target),
              size: (mediaFile as any).length || 0,
              dimensions: ImageParser.getImageDimensions(mediaFile),
            };
          }
        }
      }
    }

    return {
      url: null,
      type: "unknown",
      size: 0,
      dimensions: null,
    };
  }

  /**
   * Create R2 URL from media file buffer (async version)
   * @param mediaFile - Image file buffer
   * @param filename - Original filename
   * @param r2Storage - R2 storage binding (optional, falls back to base64)
   * @returns R2 URL, data URL, or placeholder
   */
  static async createImageUrl(
    mediaFile: Uint8Array,
    filename: string,
    r2Storage: any = null,
  ): Promise<string> {
    if (!mediaFile || !(mediaFile instanceof Uint8Array)) {
      return "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgZmlsbD0iI2RkZCIvPjx0ZXh0IHg9IjUwIiB5PSI1NSIgZm9udC1mYW1pbHk9IkFyaWFsIiBmb250LXNpemU9IjEyIiBmaWxsPSIjOTk5IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5JbWFnZTwvdGV4dD48L3N2Zz4=";
    }

    // If R2 storage is available, use it
    if (r2Storage) {
      try {
        // Generate hash for deduplication
        const hash = await ImageParser.generateImageHash(mediaFile);
        const type = ImageParser.getImageType(filename);
        const imagePath = `images/${hash}.${type}`;

        // Check if image already exists in R2
        const existing = await r2Storage.get(imagePath);
        if (existing) {
          return `/api/images/${hash}.${type}`;
        }

        // Upload to R2
        const mimeType = ImageParser.getMimeType(type);
        await r2Storage.put(imagePath, mediaFile, {
          httpMetadata: {
            contentType: mimeType,
            cacheControl: "public, max-age=31536000", // 1 year cache
          },
          customMetadata: {
            originalName: filename,
            uploadedAt: new Date().toISOString(),
            size: mediaFile.byteLength.toString(),
            hash: hash,
          },
        });

        return `/api/images/${hash}.${type}`;
      } catch (error) {
        console.error(`❌ R2 upload failed for ${filename}:`, error);
        // Fall back to base64
      }
    }

    // Fallback to base64 data URL
    return ImageParser.createDataUrl(mediaFile, filename);
  }

  /**
   * Create data URL from media file buffer (legacy method)
   * @param mediaFile - Image file buffer
   * @param filename - Original filename
   * @returns data URL or placeholder
   */
  static createDataUrl(mediaFile: Uint8Array, filename: string): string {
    if (!mediaFile || !(mediaFile instanceof Uint8Array)) {
      return "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgZmlsbD0iI2RkZCIvPjx0ZXh0IHg9IjUwIiB5PSI1NSIgZm9udC1mYW1pbHk9IkFyaWFsIiBmb250LXNpemU9IjEyIiBmaWxsPSIjOTk5IiB0ZXh0LWFuY2hvcj0ibWlkZGxlIj5JbWFnZTwvdGV4dD48L3N2Zz4=";
    }

    const type = ImageParser.getImageType(filename);
    const mimeType = ImageParser.getMimeType(type);
    const base64 = ImageParser.uint8ArrayToBase64(mediaFile);

    return `data:${mimeType};base64,${base64}`;
  }

  /**
   * Generate SHA-256 hash of image data for deduplication
   * @param uint8Array - Binary image data
   * @returns SHA-256 hash as hex string
   */
  static async generateImageHash(uint8Array: Uint8Array): Promise<string> {
    const hashBuffer = await crypto.subtle.digest("SHA-256", uint8Array);
    const hashArray = new Uint8Array(hashBuffer);
    return Array.from(hashArray)
      .map((b) => b.toString(16).padStart(2, "0"))
      .join("");
  }

  /**
   * Convert Uint8Array to base64 string (Cloudflare Workers compatible)
   * @param uint8Array - Binary data
   * @returns base64 string
   */
  static uint8ArrayToBase64(uint8Array: Uint8Array): string {
    let binary = "";
    for (let i = 0; i < uint8Array.byteLength; i++) {
      binary += String.fromCharCode(uint8Array[i]);
    }
    return btoa(binary);
  }

  /**
   * Get image type from filename
   * @param filename - Image filename
   * @returns image type
   */
  static getImageType(filename: string): string {
    const ext = filename.split(".").pop()?.toLowerCase();
    return ext || "unknown";
  }

  /**
   * Get MIME type from image type
   * @param type - Image type
   * @returns MIME type
   */
  static getMimeType(type: string): string {
    const mimeTypes: Record<string, string> = {
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      png: "image/png",
      gif: "image/gif",
      bmp: "image/bmp",
      webp: "image/webp",
      svg: "image/svg+xml",
      tiff: "image/tiff",
      tif: "image/tiff",
    };

    return mimeTypes[type] || "application/octet-stream";
  }

  /**
   * Parse image effects from blip fill
   * @param blipFill - Blip fill properties
   * @returns effects information
   */
  static parseImageEffects(
    blipFill: XMLNode | null | undefined,
  ): ImageEffectsInfo {
    const effects: ImageEffectsInfo = {
      opacity: 1,
      filter: null,
      borderRadius: 0,
      shadow: null,
      effectsList: [],
    };

    if (!blipFill) return effects;

    // Parse alpha/opacity
    const blip = BaseParser.safeGet(blipFill, "blip");
    if (blip) {
      // Look for alpha modulation
      const alphaModFix = BaseParser.asNumber(
        BaseParser.safeGet(blip, "alphaModFix.$amt"),
        0,
      );
      if (alphaModFix) {
        effects.opacity = alphaModFix / 100000; // PowerPoint uses 100000 = 100%
      }

      // Grayscale effect
      if (BaseParser.safeGet(blip, "grayscl")) {
        effects.filter = "grayscale(100%)";
        effects.effectsList.push("grayscale");
      }

      // Bi-level (black and white)
      if (BaseParser.safeGet(blip, "biLevel")) {
        effects.filter = "contrast(1000%) brightness(50%)";
        effects.effectsList.push("bilevel");
      }
    }

    return effects;
  }

  /**
   * Parse image cropping information
   * @param blipFill - Blip fill properties
   * @returns cropping information
   */
  static parseCropping(
    blipFill: XMLNode | null | undefined,
  ): ImageCroppingInfo {
    const srcRect = BaseParser.safeGet(blipFill, "srcRect.$", null) as any;
    if (!srcRect || !BaseParser.isXMLNode(srcRect)) {
      return {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
        isCropped: false,
      };
    }

    // PowerPoint uses percentages * 1000 (e.g., 10000 = 10%)
    return {
      left: BaseParser.asNumber(srcRect.l, 0) / 1000,
      top: BaseParser.asNumber(srcRect.t, 0) / 1000,
      right: BaseParser.asNumber(srcRect.r, 0) / 1000,
      bottom: BaseParser.asNumber(srcRect.b, 0) / 1000,
      isCropped: !!(srcRect.l || srcRect.t || srcRect.r || srcRect.b),
    };
  }

  /**
   * Get basic image dimensions (simplified - would need proper image parsing)
   * @param imageBuffer - Image file buffer
   * @returns dimensions
   */
  static getImageDimensions(
    imageBuffer: Uint8Array | Buffer,
  ): ImageDimensions | null {
    // This is a placeholder - in a real implementation, you'd use
    // an image parsing library like sharp or image-size
    return null;
  }
}
