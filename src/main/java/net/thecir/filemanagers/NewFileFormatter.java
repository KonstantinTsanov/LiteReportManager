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
import net.thecir.constants.Constants;
import net.thecir.enums.Platforms;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
        firstSheetTableSelloutFormat(sheet);
        firstSheetTableStockFormat(sheet);
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
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetStockLabel(Sheet sheet) {
        String labelText = "Stock";
        String cellRange = "BG1:BJ1";
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetTotalLabel(Sheet sheet) {
        String labelText = "Total";
        String cellRange = "BL1:BO1";
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetSelloutNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "A" + Constants.PLATFORMS_FIRSTROW + ":A" + Constants.PLATFORMS_TABLE_LASTROW;
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetStockNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "BG" + Constants.PLATFORMS_FIRSTROW + ":BG" + Constants.PLATFORMS_TABLE_LASTROW;
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetTotalNamcoLabel(Sheet sheet) {
        String labelText = "Namco";
        String cellRange = "BL" + Constants.PLATFORMS_FIRSTROW + ":BL" + Constants.PLATFORMS_TABLE_LASTROW;
        formatFirstSheetLabels(labelText, cellRange, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetTableSelloutFormat(Sheet sheet) {
        //Table range
        String cellRange = "B" + Constants.PLATFORMS_TABLE_FIRSTROW + ":BE" + Constants.PLATFORMS_TABLE_LASTROW;
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        //Create all cells in cellRange
        createCells(region, sheet);
        // Put THIN border around each cell in the cellRange
        XSSFCellStyle borderStyle = (XSSFCellStyle) wb.createCellStyle();
        borderCells(borderStyle, BorderStyle.THIN);
        applyStyle(region, sheet, borderStyle);
        //Format the "Total" row
        XSSFCellStyle totalBarStyle = (XSSFCellStyle) wb.createCellStyle();
        totalBarStyle.cloneStyleFrom(borderStyle);
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        totalBarStyle.setFont(font);
        applyStyle(CellRangeAddress.valueOf("B" + Constants.PLATFORMS_TABLE_LASTROW + ":BC" + Constants.PLATFORMS_TABLE_LASTROW), sheet, totalBarStyle);
        //TODO
        //The first "week" cell
//        CellReference cr = new CellReference("C" + Constants.PLATFORMS_TABLE_FIRSTROW);
        //Align the week bar to the right.
//        Cell cell = sheet.getRow(cr.getRow()).getCell(cr.getCol());
        //Set value to the first cell
        //cell.setCellValue("w0");
        XSSFCellStyle weekBarStyle = (XSSFCellStyle) wb.createCellStyle();
        weekBarStyle.cloneStyleFrom(borderStyle);
        align(weekBarStyle, HorizontalAlignment.RIGHT);
        applyStyle(CellRangeAddress.valueOf("C" + Constants.PLATFORMS_TABLE_FIRSTROW + ":BB" + Constants.PLATFORMS_TABLE_FIRSTROW), sheet, weekBarStyle);
        //Puts the platforms in their specific locations
        int row = Constants.PLATFORMS_FIRSTROW;
        for (Platforms platform : Platforms.values()) {
            CellReference leftCellRef = new CellReference("B" + row);
            CellReference RightCellRef = new CellReference("BC" + row);
            Cell leftCell = sheet.getRow(leftCellRef.getRow()).getCell(leftCellRef.getCol());
            leftCell.setCellValue(platform.toString());
            Cell rightCell = sheet.getRow(RightCellRef.getRow()).getCell(RightCellRef.getCol());
            rightCell.setCellValue(platform.toString());
            row++;
        }
        //Sets the "Total" label
        CellReference totalLeft = new CellReference("B" + Constants.PLATFORMS_TABLE_LASTROW);
        CellReference totalRight = new CellReference("BC" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell cell1 = sheet.getRow(totalLeft.getRow()).getCell(totalLeft.getCol());
        Cell cell2 = sheet.getRow(totalRight.getRow()).getCell(totalRight.getCol());
        cell1.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        cell2.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        //Sets the weekly total formulas
        for (int column = CellReference.convertColStringToIndex("C"); column <= CellReference.convertColStringToIndex("BB"); column++) {
            String letter = CellReference.convertNumToColString(column);
            CellReference totalCell = new CellReference(letter + Constants.PLATFORMS_TABLE_LASTROW);
            Cell cell3 = sheet.getRow(totalCell.getRow()).getCell(totalCell.getCol());
            cell3.setCellType(CellType.FORMULA);
            cell3.setCellFormula("SUM(" + letter + Constants.PLATFORMS_FIRSTROW + ":" + letter + Constants.PLATFORMS_LASTROW + ")");
        }
        XSSFCellStyle totalPcsStyle = (XSSFCellStyle) wb.createCellStyle();
        totalPcsStyle.cloneStyleFrom(borderStyle);
        //Same as before
        totalBarStyle.setFont(font);
        applyStyle(CellRangeAddress.valueOf("BD" + Constants.PLATFORMS_TABLE_FIRSTROW), sheet, totalPcsStyle);
        CellReference totalPcs = new CellReference("BD" + Constants.PLATFORMS_TABLE_FIRSTROW);
        Cell totalPcsLabelCell = sheet.getRow(totalPcs.getRow()).getCell(totalPcs.getCol());
        totalPcsLabelCell.setCellValue(Constants.TOTAL_PCS);
        //Sets color of the total pcs column
        XSSFCellStyle totalPcsColumn = (XSSFCellStyle) wb.createCellStyle();
        totalPcsColumn.cloneStyleFrom(borderStyle);
        setBackgroundColor(totalPcsColumn, 0x92D050);
        applyStyle(CellRangeAddress.valueOf("BD" + Constants.PLATFORMS_FIRSTROW + ":BD" + Constants.PLATFORMS_LASTROW), sheet, totalPcsColumn);
        //From platforms firstrow to the last row of the table, so i can sum all weeks
        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_TABLE_LASTROW; i++) {
            CellReference totalPcsCellRef = new CellReference("BD" + i);
            Cell totalPcsCell = sheet.getRow(totalPcsCellRef.getRow()).getCell(totalPcsCellRef.getCol());
            totalPcsCell.setCellType(CellType.FORMULA);
            totalPcsCell.setCellFormula("SUM(C" + i + ":BB" + i + ")");
        }

        XSSFCellStyle percentageStyle = (XSSFCellStyle) wb.createCellStyle();
        percentageStyle.cloneStyleFrom(borderStyle);
        percentageStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        applyStyle(CellRangeAddress.valueOf("BE" + Constants.PLATFORMS_FIRSTROW + ":BE" + Constants.PLATFORMS_TABLE_LASTROW), sheet, percentageStyle);

        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_LASTROW; i++) {
            CellReference percentageCellsRef = new CellReference("BE" + i);
            Cell percentageCell = sheet.getRow(percentageCellsRef.getRow()).getCell(percentageCellsRef.getCol());
            percentageCell.setCellType(CellType.FORMULA);
            percentageCell.setCellFormula("BD" + i + "/BD" + Constants.PLATFORMS_TABLE_LASTROW);
        }
        CellReference totalPercentageCellRef = new CellReference("BE" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalPercentageCell = sheet.getRow(totalPercentageCellRef.getRow()).getCell(totalPercentageCellRef.getCol());
        totalPercentageCell.setCellType(CellType.FORMULA);
        totalPercentageCell.setCellFormula("SUM(BE" + Constants.PLATFORMS_FIRSTROW + ":BE" + Constants.PLATFORMS_LASTROW + ")");
    }

    private void firstSheetTableStockFormat(Sheet sheet) {
        String cellRange = "BH" + Constants.PLATFORMS_FIRSTROW + ":" + "BJ" + Constants.PLATFORMS_TABLE_LASTROW;
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        createCells(region, sheet);
        XSSFCellStyle borderStyle = (XSSFCellStyle) wb.createCellStyle();
        borderCells(borderStyle, BorderStyle.THIN);
        applyStyle(region, sheet, borderStyle);

        int row = Constants.PLATFORMS_FIRSTROW;
        for (Platforms platform : Platforms.values()) {
            CellReference cellRef = new CellReference("BH" + row);
            Cell cell = sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
            cell.setCellValue(platform.toString());
            row++;
        }
        CellReference totalCellRef = new CellReference("BH" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalCell = sheet.getRow(totalCellRef.getRow()).getCell(totalCellRef.getCol());
        totalCell.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);

        CellReference totalCellFormulaRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalCellFormula = sheet.getRow(totalCellFormulaRef.getRow()).getCell(totalCellFormulaRef.getCol());
        totalCellFormula.setCellType(CellType.FORMULA);
        totalCellFormula.setCellFormula("SUM(BI" + Constants.PLATFORMS_FIRSTROW + ":BI" + Constants.PLATFORMS_LASTROW + ")");

        CellReference daysInStockCellRef = new CellReference("BJ" + Constants.PLATFORMS_TABLE_FIRSTROW);
        Cell daysInStockCell = sheet.getRow(daysInStockCellRef.getRow()).getCell(daysInStockCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        daysInStockCell.setCellValue(Constants.DAYS_IN_STOCK);

        XSSFCellStyle percentageStyle = (XSSFCellStyle) wb.createCellStyle();
        percentageStyle.cloneStyleFrom(borderStyle);
        percentageStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        applyStyle(CellRangeAddress.valueOf("BJ" + Constants.PLATFORMS_FIRSTROW + ":BJ" + Constants.PLATFORMS_TABLE_LASTROW), sheet, percentageStyle);
    }

    //--------------END OF FIRST SHEET FORMATTERS-----------------//
    //--------------TOOLS-----------------
    private void formatFirstSheetLabels(String labelText, String cellRange, Sheet sheet, int backgroundColor, short rotationDegrees) {
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        borderCells(style, BorderStyle.THIN);
        align(style, HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        setRotation(style, rotationDegrees);
        setBackgroundColor(style, backgroundColor);
        createCells(region, sheet);
        applyStyle(region, sheet, style);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    /**
     * Creates a border around the selected cells
     *
     * @param style Style to apply the border to.
     * @param borderStyle Type of border.
     */
    private void borderCells(XSSFCellStyle style, BorderStyle borderStyle) {
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
    private void align(XSSFCellStyle style, HorizontalAlignment alignment) {
        style.setAlignment(alignment);
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
            Row row = sheet.getRow(i);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cell = CellUtil.getCell(row, j);
                cell.setCellStyle(style);
            }
        }
    }

    private void createCells(CellRangeAddress region, Sheet sheet) {
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row row = CellUtil.getRow(i, sheet);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                row.createCell(j);
            }
        }
    }

//--------------END OF TOOLS-----------------
}
