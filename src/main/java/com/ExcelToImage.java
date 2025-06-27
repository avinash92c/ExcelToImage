package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;

public class ExcelToImage {

    public static void main(String[] args) throws Exception {
//        if (args.length == 0) {
//            System.out.println("Usage: java ExcelToImageConverter <input-file.xlsm>");
//            return;
//        }

        FileInputStream fis = new FileInputStream("invoice-template.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            if (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)) continue;

            XSSFSheet sheet = workbook.getSheetAt(i);
            BufferedImage image = renderSheetToImage(sheet,workbook);
            File out = new File("sheet_" + i + "_" + sheet.getSheetName() + ".png");
            ImageIO.write(image, "png", out);
            System.out.println("Saved: " + out.getName());
        }

        workbook.close();
        fis.close();
    }

    private static BufferedImage renderSheetToImage(XSSFSheet sheet, XSSFWorkbook workbook) {
        int rowHeightPx = 20;
        int colWidthPx = 100;
        int margin = 10;

        int visibleRowCount = 0;
        int visibleColCount = 0;

        for (Row row : sheet) {
            if (row.getZeroHeight()) continue;
            visibleRowCount++;
            for (Cell cell : row) {
                int colIdx = cell.getColumnIndex();
                if (sheet.isColumnHidden(colIdx)) continue;
                visibleColCount = Math.max(visibleColCount, colIdx + 1);
            }
        }

        int width = margin * 2 + visibleColCount * colWidthPx;
        int height = margin * 2 + visibleRowCount * rowHeightPx;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2 = image.createGraphics();
        g2.setColor(Color.WHITE);
        g2.fillRect(0, 0, width, height);
        g2.setFont(new Font("Arial", Font.PLAIN, 12));

        int y = margin;
        for (Row row : sheet) {
            if (row.getZeroHeight()) continue;
            int x = margin;
            for (int c = 0; c < visibleColCount; c++) {
                if (sheet.isColumnHidden(c)) continue;

                Cell cell = row.getCell(c);
                String text = (cell != null) ? getCellText(cell) : "";

                g2.setColor(Color.WHITE);
                g2.fillRect(x, y, colWidthPx, rowHeightPx);

                if (cell != null) {
                    XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
                    XSSFColor bgColor = style.getFillForegroundXSSFColor();
                    if (bgColor != null && bgColor.getRGB() != null) {
                        byte[] rgb = bgColor.getRGB();
                        g2.setColor(new Color(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                        g2.fillRect(x, y, colWidthPx, rowHeightPx);
                    }

                    g2.setColor(Color.GRAY);
                    g2.drawRect(x, y, colWidthPx, rowHeightPx);

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
                    g2.drawString(text, x + 5, y + 15);
                }

                x += colWidthPx;
            }
            y += rowHeightPx;
        }

        g2.dispose();
        return image;
    }

    private static String getCellText(Cell cell) {
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
