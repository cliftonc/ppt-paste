/**
 * Video component parser for PowerPoint shapes containing embedded videos
 */

import { BaseParser } from "./BaseParser.js";
import {
  XMLNode,
  VideoComponent,
  NormalizedVideoComponent,
} from "../types/index.js";

export class VideoParser extends BaseParser {
  /**
   * Parse video component from normalized data (works for both PPTX and clipboard)
   * @param videoComponent - Normalized video component
   * @param relationships - Relationship data
   * @param mediaFiles - Media files
   * @param componentIndex - Component index
   * @param slideIndex - Slide index (position in presentation order)
   * @param r2Storage - Optional R2 storage for thumbnail hosting
   * @param slideRelsFile - The slide's own .rels part; preferred over deriving one from
   *   slideIndex, which assumes presentation order matches the slide's FILE number
   * @returns Parsed video component
   */
  static async parseFromNormalized(
    videoComponent: NormalizedVideoComponent,
    relationships: Record<string, any>,
    mediaFiles: Record<string, Uint8Array>,
    componentIndex: number,
    slideIndex: number,
    zIndex: number,
    r2Storage: any = null,
    slideRelsFile: string | null = null
  ): Promise<VideoComponent | null> {
    const { data, spPr, nvPicPr, blipFill, namespace, relationshipId } =
      videoComponent;

    if (!spPr || !nvPicPr) {
      throw new Error("No spPr or nvPicPr found in normalized video component");
    }

    // Extract positioning from spPr (namespaces already stripped)
    const xfrm = BaseParser.getNode(spPr, "xfrm");
    const transform = VideoParser.parseTransform(xfrm);

    // Extract component info from nvPicPr
    const cNvPr = BaseParser.getNode(nvPicPr, "cNvPr");
    const originalName = BaseParser.getString(cNvPr, "$name", "");
    const componentName = BaseParser.generateComponentId('video', componentIndex, originalName);
    const description = BaseParser.getString(cNvPr, "$descr", "");

    // Extract video reference from nvPr
    const nvPr = BaseParser.getNode(nvPicPr, "nvPr");
    const videoFile = BaseParser.getNode(nvPr, "videoFile");
    const videoRelationshipId =
      relationshipId || BaseParser.getString(videoFile, "$link", "");

    // Find the video URL using relationships
    let videoUrl: string | null = null;
    let embedType: "youtube" | "vimeo" | "generic" = "generic";

    // Detect if this is clipboard format vs PPTX format
    const isClipboardFormat = Object.keys(relationships || {}).some((key) =>
      key.includes("clipboard/"),
    );

    if (videoRelationshipId && relationships) {
      if (!isClipboardFormat && (slideRelsFile || slideIndex !== null)) {
        // For PPTX files, look in the slide's relationship file
        const currentSlideRelFile =
          slideRelsFile || `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`;

        if (relationships[currentSlideRelFile]) {
          const relsData =
            BaseParser.safeGet(
              relationships[currentSlideRelFile],
              "Relationships.Relationship",
            ) || [];
          const rels = Array.isArray(relsData) ? relsData : [relsData];
          const rel = rels.find((r: any) => r && r.$Id === videoRelationshipId);

          if (rel) {
            videoUrl = rel.$Target;

            // Determine embed type based on URL
            if (videoUrl) {
              if (
                videoUrl.includes("youtube.com") ||
                videoUrl.includes("youtu.be")
              ) {
                embedType = "youtube";
              } else if (videoUrl.includes("vimeo.com")) {
                embedType = "vimeo";
              }
            }
          }
        }
      } else {
        // Clipboard format - look directly in relationships
        const slideRelationships = relationships[slideIndex];
        if (slideRelationships && slideRelationships[videoRelationshipId]) {
          const relationship = slideRelationships[videoRelationshipId];
          videoUrl = relationship.Target || relationship.target;

          // Determine embed type based on URL
          if (videoUrl) {
            if (
              videoUrl.includes("youtube.com") ||
              videoUrl.includes("youtu.be")
            ) {
              embedType = "youtube";
            } else if (videoUrl.includes("vimeo.com")) {
              embedType = "vimeo";
            }
          }
        }
      }
    }

    // Extract thumbnail image reference from blipFill if available
    let thumbnailSrc: string | null = null;
    if (blipFill) {
      const blip = BaseParser.safeGet(blipFill, "blip");
      const thumbnailRelationshipId = BaseParser.safeGet(blip, "embed");

      if (thumbnailRelationshipId && relationships && mediaFiles) {
        let thumbnailPath: string | null = null;

        if (!isClipboardFormat && (slideRelsFile || slideIndex !== null)) {
          // For PPTX files, look in the slide's relationship file
          const currentSlideRelFile =
            slideRelsFile || `ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`;

          if (relationships[currentSlideRelFile]) {
            const relsData =
              BaseParser.safeGet(
                relationships[currentSlideRelFile],
                "Relationships.Relationship",
              ) || [];
            const rels = Array.isArray(relsData) ? relsData : [relsData];
            const rel = rels.find(
              (r: any) => r && r.$Id === thumbnailRelationshipId,
            );

            if (rel) {
              thumbnailPath = rel.$Target;
              // Convert relative path to absolute
              if (thumbnailPath?.startsWith("../")) {
                thumbnailPath = `ppt/${thumbnailPath.slice(3)}`;
              } else if (thumbnailPath && !thumbnailPath.startsWith("ppt/")) {
                thumbnailPath = thumbnailPath.startsWith("media/")
                  ? `ppt/${thumbnailPath}`
                  : thumbnailPath;
              }
            }
          }
        } else {
          // Clipboard format - look directly in relationships
          const slideRelationships = relationships[slideIndex];
          if (
            slideRelationships &&
            slideRelationships[thumbnailRelationshipId]
          ) {
            const thumbnailRelationship =
              slideRelationships[thumbnailRelationshipId];
            thumbnailPath =
              thumbnailRelationship.Target || thumbnailRelationship.target;
          }
        }

        if (thumbnailPath) {
          // Try to find the thumbnail in media files
          const mediaKey = Object.keys(mediaFiles).find(
            (key) =>
              key === thumbnailPath ||
              key.includes(thumbnailPath!) ||
              thumbnailPath!.includes(key.split("/").pop() || ""),
          );

          if (mediaKey && mediaFiles[mediaKey]) {
            const mediaData = mediaFiles[mediaKey];
            // Use ImageParser to handle R2 upload with same pattern as regular images
            const { ImageParser } = await import("./ImageParser.js");
            thumbnailSrc = await ImageParser.createImageUrl(
              mediaData,
              thumbnailPath,
              r2Storage,
            );
          }
        }
      }
    }

    return {
      id: componentName,
      type: "video",
      content: description || componentName,
      x: transform.x,
      y: transform.y,
      width: transform.width,
      height: transform.height,
      slideIndex,
      zIndex,
      style: {
        rotation: transform.rotation || 0,
      },
      url: videoUrl || "",
      thumbnailSrc: thumbnailSrc || "",
      title: description || componentName,
      embedType,
      metadata: {
        namespace,
        name: componentName,
        description,
        videoRelationshipId,
        thumbnailRelationshipId: blipFill
          ? BaseParser.safeGet(BaseParser.safeGet(blipFill, "blip"), "embed")
          : null,
        originalFormat: "normalized",
        videoUrl: videoUrl,
      },
    };
  }

  /**
   * Get MIME type for image files
   */
  private static getImageMimeType(filePath: string): string {
    const extension = filePath.toLowerCase().split(".").pop();
    switch (extension) {
      case "jpg":
      case "jpeg":
        return "image/jpeg";
      case "png":
        return "image/png";
      case "gif":
        return "image/gif";
      case "webp":
        return "image/webp";
      default:
        return "image/jpeg"; // Default fallback
    }
  }

  /**
   * Extract video information from relationships
   * This is a helper method for future enhancements
   */
  static extractVideoInfo(url: string): {
    embedType: "youtube" | "vimeo" | "generic";
    videoId?: string;
  } {
    if (!url) return { embedType: "generic" };

    // YouTube URL patterns
    const youtubePatterns = [
      /(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^"&?\/\s]{11})/,
      /youtube\.com\/embed\/([^"&?\/\s]{11})/,
    ];

    for (const pattern of youtubePatterns) {
      const match = url.match(pattern);
      if (match) {
        return { embedType: "youtube", videoId: match[1] };
      }
    }

    // Vimeo URL patterns
    const vimeoPattern = /vimeo\.com\/(?:.*\/)?(\d+)/;
    const vimeoMatch = url.match(vimeoPattern);
    if (vimeoMatch) {
      return { embedType: "vimeo", videoId: vimeoMatch[1] };
    }

    return { embedType: "generic" };
  }
}
