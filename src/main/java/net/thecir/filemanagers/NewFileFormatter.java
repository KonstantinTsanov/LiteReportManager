/*
 * The MIT License
 *
 * Copyright 2017 Konstantin Tsanov <k.tsanov@gmail.com>.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */
package net.thecir.filemanagers;

import java.awt.Color;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
public class NewFileFormatter {

    private final Workbook wb;

    public NewFileFormatter(Workbook workbook) {
        this.wb = workbook;
    }

    public void formatWorkbook() {
        addAndFormatSheets();
    }

    /**
     * Adds all necessary sheets to the workbook.
     */
    private void addAndFormatSheets() {
        Sheet sheet1 = wb.createSheet("Review");
        Sheet sheet2 = wb.createSheet("Top 5 statistics");
        Sheet sheet3 = wb.createSheet("Overall sales by platform");
        Sheet sheet4 = wb.createSheet("Overall sales by game");
        formatFirstSheet(sheet1);
        formatSecondSheet(sheet2);
        formatThirdSheet(sheet3);
        formatFourthSheet(sheet4);
    }

    private void formatFirstSheet(Sheet sheet) {
        firstSheetSelloutLabel(sheet);
        firstSheetStockLabel(sheet);
        firstSheetTotalLabel(sheet);
        firstSheetSelloutNamcoLabel(sheet);
        firstSheetStockNamcoLabel(sheet);
        firstSheetTotalNamcoLabel(sheet);
    }

    private void formatSecondSheet(Sheet sheet) {

    }

    private void formatThirdSheet(Sheet sheet) {

    }

    private void formatFourthSheet(Sheet sheet) {

    }

    //--------------FIRST SHEET FORMATTERS-----------------//
    private void firstSheetSelloutLabel(Sheet sheet) {
        String labelText = "Sell out";
        String cellRange = "A1:BE1";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDAEEF3, (short) 0);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    private void firstSheetStockLabel(Sheet sheet) {
        String labelText = "Stock";
        String cellRange = "BG1:BJ1";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDAEEF3, (short) 0);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    private void firstSheetTotalLabel(Sheet sheet) {
        String labelText = "Total";
        String cellRange = "BL1:BO1";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDAEEF3, (short) 0);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    private void firstSheetSelloutNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "A4:A16";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDA9694, (short) 90);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    private void firstSheetStockNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "BG4:BG16";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDA9694, (short) 90);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    private void firstSheetTotalNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "BL4:BL16";
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        labelsFormat(region, sheet, 0xDA9694, (short) 90);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    //--------------END OF FIRST SHEET FORMATTERS-----------------//
    //--------------TOOLS-----------------
    private void labelsFormat(CellRangeAddress region, Sheet sheet, int backgroundColor, short rotationDegrees) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        border(style, BorderStyle.THIN);
        center(style);
        setBold(style);
        setRotation(style, rotationDegrees);
        setBackgroundColor(style, backgroundColor);
        applyStyle(region, sheet, style);
    }

    /**
     * Creates a border around the selected cells
     *
     * @param style Style to apply the border to.
     * @param borderStyle Type of border.
     */
    private void border(XSSFCellStyle style, BorderStyle borderStyle) {
        style.setBorderBottom(borderStyle);
        style.setBorderLeft(borderStyle);
        style.setBorderRight(borderStyle);
        style.setBorderTop(borderStyle);
    }

    /**
     * Applies centering to a style.
     *
     * @param style Style to apply center to.
     */
    private void center(XSSFCellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
    }

    /**
     * Applies bold effect to a style
     *
     * @param style Style to apply bold to.
     */
    private void setBold(XSSFCellStyle style) {
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
    }

    private void setRotation(XSSFCellStyle style, short degrees) {
        style.setRotation(degrees);
    }

    /**
     * Sets background color to selected style.
     *
     * @param style Style to apply color to.
     * @param color Color to be applied.
     */
    private void setBackgroundColor(XSSFCellStyle style, int color) {
        XSSFColor myColor = new XSSFColor(new Color(color));
        style.setFillForegroundColor(myColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    /**
     * Merges the selected cells within a sheet.
     *
     * @param region - selected cells to be merged.
     * @param sheet - sheet, on which the cells are located.
     */
    private void mergeCells(CellRangeAddress region, Sheet sheet) {
        sheet.addMergedRegion(region);
    }

    /**
     * Set value to the first cell in a set of merged cells.
     *
     * @param region - region of merged cells.
     * @param sheet - sheet where the region is located.
     * @param value - value to be set.
     */
    private void setMergedCellsValue(CellRangeAddress region, Sheet sheet, String value) {
        Row row = sheet.getRow(region.getFirstRow());
        Cell cell = row.getCell(region.getFirstColumn());
        cell.setCellValue(value);
    }

    /**
     * Iterates over a region of cells and applies style to them.
     *
     * @param region - region of cells.
     * @param sheet - sheet where the region is located.
     * @param style - style to be applied.
     */
    private void applyStyle(CellRangeAddress region, Sheet sheet, XSSFCellStyle style) {
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row row = CellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellStyle(style);
            }
        }
    }

//--------------END OF TOOLS-----------------
}
