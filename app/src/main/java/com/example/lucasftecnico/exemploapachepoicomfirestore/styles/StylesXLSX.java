package com.example.lucasftecnico.exemploapachepoicomfirestore.styles;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.HashMap;
import java.util.Map;

public class StylesXLSX {


    private static void createBorder(CellStyle cellStyle) {
        short borderColor = IndexedColors.GREY_50_PERCENT.getIndex();

        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setTopBorderColor(borderColor);

        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setRightBorderColor(borderColor);

        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBottomBorderColor(borderColor);

        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor(borderColor);
    }


    public static Map<String, CellStyle> createStyles(Workbook workbook, XSSFSheet xssfSheet) {
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;

        // Style título
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 48);
        titleFont.setColor(IndexedColors.GREEN.getIndex());
        titleFont.setBold(true);
        style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.BIG_SPOTS);
        style.setFont(titleFont);
        createBorder(style);
        styles.put("titulo", style);

        // Style para subtítulo
        Font fontSubtitulo = workbook.createFont();
        fontSubtitulo.setFontHeightInPoints((short) 15);
        fontSubtitulo.setColor(IndexedColors.WHITE.getIndex());
        fontSubtitulo.setBold(true);
        style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.DARK_GREEN.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(fontSubtitulo);
        createBorder(style);
        styles.put("subtitulo", style);

        // Style para corpo
        Font fontCorpo01 = workbook.createFont();
        fontCorpo01.setFontHeightInPoints((short) 12);
        fontCorpo01.setColor(IndexedColors.AUTOMATIC.getIndex());
        style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(fontCorpo01);
        createBorder(style);
        styles.put("corpo_01", style);

        // Style para corpo
        Font fontCorpo02 = workbook.createFont();
        fontCorpo02.setFontHeightInPoints((short) 12);
        fontCorpo02.setColor(IndexedColors.AUTOMATIC.getIndex());
        style = workbook.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(fontCorpo02);
        createBorder(style);
        styles.put("corpo_02", style);

        return styles;
    }
}
