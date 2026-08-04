/**
 * Table component parser for PowerPoint tables
 */

import { BaseParser } from "./BaseParser.js";
import { TextParser } from "./TextParser.js";
import {
  XMLNode,
  TableComponent,
  NormalizedTableComponent,
  TableDimensions,
  TipTapDocumentNode,
  TipTapTableCellNode,
  TipTapTableRowNode,
  TipTapTableNode,
  TipTapParagraphNode,
} from "../types/index.js";

export class TableParser extends BaseParser {
  /**
   * Parse table component from normalized data
   * @param tableComponent - Normalized table component
   * @param componentIndex - Component index
   * @param slideIndex - Slide index
   * @returns Parsed table component
   */
  static async parseFromNormalized(
    tableComponent: NormalizedTableComponent,
    componentIndex: number,
    slideIndex: number,
    zIndex: number
  ): Promise<TableComponent | null> {
    const { graphicData, spPr, nvGraphicFramePr, namespace } = tableComponent;

    if (!graphicData) {
      throw new Error("No graphicData found in normalized table component");
    }

    // Extract table data from tbl element
    const tableData = this.extractTableData(graphicData);
    if (!tableData || tableData.length === 0) {
      return null;
    }

    // Create richText structure matching client-side implementation
    const richText = this.createTableRichText(tableData, true); // Assume first row is header

    // Extract positioning from spPr (which should be the xfrm from graphicFrame)
    const transform = this.parseTransform(spPr); // spPr is actually xfrm for graphicFrame

    // Extract component info from nvGraphicFramePr
    const originalName = BaseParser.getString(nvGraphicFramePr, "cNvPr.$name", "");
    const componentName = BaseParser.generateComponentId('table', componentIndex, originalName);

    // Calculate table dimensions
    const { rows, cols } = this.getTableDimensions(tableData);

    return {
      id: componentName,
      slideIndex,
      zIndex,
      type: "table",
      content: `Table (${rows} rows × ${cols} columns)`,
      x: transform.x,
      y: transform.y,
      width: transform.width,
      height: transform.height,
      style: {
        rotation: transform.rotation,
        borderColor: "#ffffff",
        fillColor: "transparent",
      },
      rows: tableData.map((rowData) => ({
        cells: rowData.map((cellContent) => ({
          content: cellContent || "",
          style: {},
        })),
      })),
      columns: cols,
      metadata: {
        tableData: tableData,
        rows: rows,
        cols: cols,
        hasHeader: true,
        source: "server-parsed",
        format: namespace || "unknown",
        richText: richText, // Add richText property for TLDraw compatibility
      },
    };
  }

  /**
   * Extract table data from graphicData
   * @param graphicData - PowerPoint table graphic data
   * @returns 2D array of table cell text
   */
  static extractTableData(graphicData: XMLNode): string[][] {
    try {
      // The table structure is: graphicData -> tbl (namespace stripped already)
    const table = BaseParser.getNode(graphicData, "tbl");
    if (!table) {
      console.warn("No table element found in graphicData");
      return [];
    }

    const rowsArray = BaseParser.getArray(table, "tr", []);
    const tableData: string[][] = [];

    for (const row of rowsArray) {
      const cellsArray = BaseParser.getArray(row, "tc", []);
      const rowData: string[] = [];

      for (const cell of cellsArray) {
        const txBody = BaseParser.getNode(cell, "txBody");
        const cellText = this.extractCellText(txBody);
        rowData.push(cellText);
      }

      if (rowData.length > 0) {
        tableData.push(rowData);
      }
    }

    return tableData;
    } catch (error) {
      console.error("Error extracting table data:", error);
      return [];
    }
  }

  /**
   * Extract text content from a table cell
   * @param txBody - Text body element
   * @returns Cell text content
   */
  static extractCellText(txBody: XMLNode | null | undefined): string {
    if (!txBody) return "";
    try {
      // Reuse TextParser logic to ensure consistent bullet / spacing handling (bullets ignored inside tables for now).
      // Imported statically: this package is ESM, where `require` is not defined — and the
      // ReferenceError landed in the catch below, so every cell in every table came back
      // empty and said nothing about it.
      const text = TextParser.extractTextContent(txBody);
      return (text || '').trim();
    } catch (error) {
      console.error("Error extracting cell text:", error);
      return "";
    }
  }

  /**
   * Create TipTap richText table structure matching client-side implementation
   * @param tableData - 2D array of table cell text
   * @param hasHeader - Whether first row is header
   * @returns TipTap richText structure
   */
  static createTableRichText(
    tableData: string[][],
    hasHeader: boolean = true,
  ): TipTapDocumentNode {
    if (!tableData || tableData.length === 0) {
      return {
        type: "doc",
        content: [
          {
            type: "paragraph",
            content: [{ type: "text", text: "Empty table" }],
          },
        ],
      };
    }

    const tableRows: TipTapTableRowNode[] = tableData.map(
      (rowData, rowIndex) => {
        const isHeaderRow = hasHeader && rowIndex === 0;
        const cellType: "tableHeader" | "tableCell" = isHeaderRow
          ? "tableHeader"
          : "tableCell";

        const cells: TipTapTableCellNode[] = rowData.map((cellContent) => {
          const textContent =
            cellContent && cellContent.trim() ? cellContent : " ";
          return {
            type: cellType,
            attrs: {
              colspan: 1,
              rowspan: 1,
              colwidth: null,
            },
            content: [
              {
                type: "paragraph",
                attrs: {
                  dir: "auto",
                },
                content: [
                  {
                    type: "text",
                    text: textContent,
                  },
                ],
              },
            ],
          };
        });

        return {
          type: "tableRow",
          content: cells,
        };
      },
    );

    const tableNode: TipTapTableNode = {
      type: "table",
      content: tableRows,
    };

    return {
      type: "doc",
      content: [tableNode],
    };
  }

  /**
   * Get table dimensions
   * @param tableData - Table data
   * @returns {rows, cols}
   */
  static getTableDimensions(tableData: string[][]): TableDimensions {
    if (!tableData || tableData.length === 0) {
      return { rows: 0, cols: 0 };
    }
    return {
      rows: tableData.length,
      cols: tableData[0]?.length || 0,
    };
  }
}
