package com;

import java.awt.Color;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;

import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class ExcelToImageConverter {

    public static void main(String[] args) throws Exception {
//        if (args.length == 0) {
//            System.out.println("Usage: java ExcelToImageConverter <input-file.xlsm>");
//            return;
//        }

//    	FileInputStream fis = new FileInputStream("invoice-template.xlsx");
    	FileInputStream fis = new FileInputStream("Download-Sample-File-xlsm.xlsm");
    	
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) continue;

            XSSFSheet sheet = workbook.getSheetAt(i);
            BufferedImage image = renderSheetToImage(sheet, workbook);
            File out = new File("sheet_" + i + "_" + sheet.getSheetName() + ".png");
            ImageIO.write(image, "png", out);
            System.out.println("Saved: " + out.getName());
        }

        workbook.close();
        fis.close();
    }

    private static BufferedImage renderSheetToImage(XSSFSheet sheet, XSSFWorkbook workbook) {
        int margin = 10;

        // Determine visible rows and columns
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        // Calculate visible columns (max column index of all visible rows)
        int maxColIndex = 0;
        for (int r = firstRowNum; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null || row.getZeroHeight()) continue;
            for (Cell cell : row) {
                int col = cell.getColumnIndex();
                if (sheet.isColumnHidden(col)) continue;
                if (col > maxColIndex) maxColIndex = col;
            }
        }

        // Compute widths of columns in pixels and total width
        int[] colWidths = new int[maxColIndex + 1];
        int totalWidth = 0;
        for (int c = 0; c <= maxColIndex; c++) {
            if (sheet.isColumnHidden(c)) {
                colWidths[c] = 0;
            } else {
                colWidths[c] = (int) (sheet.getColumnWidth(c) * 0.075); // approximate px conversion
                totalWidth += colWidths[c];
            }
        }

        // Compute heights of rows in pixels and total height
        int[] rowHeights = new int[lastRowNum - firstRowNum + 1];
        int totalHeight = 0;
        for (int r = firstRowNum; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null || row.getZeroHeight()) {
                rowHeights[r - firstRowNum] = 0;
                continue;
            }
            int height = (int) (row.getHeightInPoints() * 1.33); // approximate px conversion
            rowHeights[r - firstRowNum] = height > 0 ? height : 20;
            totalHeight += rowHeights[r - firstRowNum];
        }

        int width = margin * 2 + totalWidth;
        int height = margin * 2 + totalHeight;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();

        // Enable antialiasing for better text quality
        g2.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

        // Fill background white
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, width, height);

        // Prepare merged cell info
        Set<String> mergedCellsDrawn = new HashSet<>();
        for (int r = firstRowNum; r <= lastRowNum; r++) {
            Row row = sheet.getRow(r);
            if (row == null || row.getZeroHeight()) continue;

            int y = margin + sumArray(rowHeights, 0, r - firstRowNum);
            int x = margin;

            for (int c = 0; c <= maxColIndex; c++) {
                if (sheet.isColumnHidden(c)) continue;

                // Check if cell is part of a merged region already drawn
                Cell cell = row.getCell(c);
                if (cell == null) {
                    x += colWidths[c];
                    continue;
                }

                CellRangeAddress mergedRegion = getMergedRegion(sheet, r, c);
                if (mergedRegion != null) {
                    // Draw merged region once at top-left cell
                    String mergedKey = mergedRegion.formatAsString();
                    if (mergedCellsDrawn.contains(mergedKey)) {
                        x += colWidths[c];
                        continue;
                    }
                    mergedCellsDrawn.add(mergedKey);

                    int startRow = mergedRegion.getFirstRow();
                    int endRow = mergedRegion.getLastRow();
                    int startCol = mergedRegion.getFirstColumn();
                    int endCol = mergedRegion.getLastColumn();

                    int mergedX = margin + sumArray(colWidths, 0, startCol);
                    int mergedY = margin + sumArray(rowHeights, 0, startRow - firstRowNum);
                    int mergedWidth = sumArray(colWidths, startCol, endCol + 1);
                    int mergedHeight = sumArray(rowHeights, startRow - firstRowNum, endRow - firstRowNum + 1);

                    drawCell(g2, sheet, workbook, sheet.getRow(startRow).getCell(startCol), mergedX, mergedY, mergedWidth, mergedHeight);

                    x += colWidths[c];
                } else {
                    // Normal single cell
                    int cellWidth = colWidths[c];
                    int cellHeight = rowHeights[r - firstRowNum];

                    drawCell(g2, sheet, workbook, cell, x, y, cellWidth, cellHeight);

                    x += cellWidth;
                }
            }
        }

        g2.dispose();
        return image;
    }

    private static void drawCell(Graphics2D g2, XSSFSheet sheet, XSSFWorkbook workbook, Cell cell,
                                 int x, int y, int width, int height) {
        // Draw background
        g2.setColor(Color.WHITE);
        g2.fillRect(x, y, width, height);

        XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
        XSSFColor bgColor = style.getFillForegroundXSSFColor();

        if (bgColor != null && bgColor.getRGB() != null) {
            byte[] rgb = bgColor.getRGB();
            Color bg = new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF);
            g2.setColor(bg);
            g2.fillRect(x, y, width, height);
        }

        // Draw border
        g2.setColor(Color.GRAY);
        g2.drawRect(x, y, width, height);

        // Draw text
        XSSFFont font = workbook.getFontAt(style.getFontIndexAsInt());
        XSSFColor fontColor = font.getXSSFColor();
        if (fontColor != null && fontColor.getRGB() != null) {
            byte[] fRgb = fontColor.getRGB();
            g2.setColor(new Color(fRgb[0] & 0xFF, fRgb[1] & 0xFF, fRgb[2] & 0xFF));
        } else {
            g2.setColor(Color.BLACK);
        }

        Font awtFont = new Font("Arial", font.getBold() ? Font.BOLD : Font.PLAIN, 12);
        g2.setFont(awtFont);

        // Text alignment (left/center/right)
        HorizontalAlignment alignment = style.getAlignment();
        FontMetrics fm = g2.getFontMetrics();

        String text = getCellText(cell);
        int textWidth = fm.stringWidth(text);
        int textHeight = fm.getAscent();

        int tx;
        switch (alignment) {
            case CENTER:
                tx = x + (width - textWidth) / 2;
                break;
            case RIGHT:
                tx = x + width - textWidth - 5;
                break;
            default: // LEFT or general
                tx = x + 5;
        }

        int ty = y + (height + textHeight) / 2 - 3; // vertically center-ish

        g2.drawString(text, tx, ty);
    }

    private static CellRangeAddress getMergedRegion(XSSFSheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.isInRange(row, col)) return region;
        }
        return null;
    }

    // Sum elements in arr from start inclusive to end exclusive
    private static int sumArray(int[] arr, int start, int end) {
        int sum = 0;
        for (int i = start; i < end && i < arr.length; i++) {
            sum += arr[i];
        }
        return sum;
    }

    private static String getCellText(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }
}
