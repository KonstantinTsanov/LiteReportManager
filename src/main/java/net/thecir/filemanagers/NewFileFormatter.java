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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
public class NewFileFormatter {

    private final XSSFWorkbook wb;

    /**
     * Creates a formatter for the new workbook
     *
     * @param workbook - workbook to be formatted.
     */
    public NewFileFormatter(XSSFWorkbook workbook) {
        this.wb = workbook;
    }

    /**
     * Formats the whole workbook
     */
    public void formatWorkbook() {
        addAndFormatSheets();
    }

    /**
     * Adds all necessary sheets to the workbook, then calls format method on
     * each.
     */
    private void addAndFormatSheets() {
        //Sheet 1 added to the default template
        XSSFSheet sheet2 = wb.createSheet(Constants.SHEET_2_NAME);
        XSSFSheet sheet3 = wb.createSheet(Constants.SHEET_3_NAME);
        XSSFSheet sheet4 = wb.createSheet(Constants.SHEET_4_NAME);
        //Get default sheet
        formatFirstSheet(wb.getSheetAt(0));
        formatSecondSheet(sheet2);
        formatThirdSheet(sheet3);
        formatFourthSheet(sheet4);
    }

    /**
     * Formats the first sheet, calling a variety of different methods,
     * formatting different areas of the sheet
     *
     * @param sheet - the first sheet must be referenced here.
     */
    private void formatFirstSheet(XSSFSheet sheet) {
        firstSheetSelloutLabel(sheet);
        firstSheetStockLabel(sheet);
        firstSheetTotalLabel(sheet);
        firstSheetSelloutNamcoLabel(sheet);
        firstSheetStockNamcoLabel(sheet);
        firstSheetTotalNamcoLabel(sheet);
        firstSheetTableSellout(sheet);
        firstSheetTableStockFormat(sheet);
        firstSheetTableTotalFormat(sheet);
    }

    /**
     * Formats the second sheet, calling a variety of different methods,
     * formatting different areas of the sheet
     *
     * @param sheet - the second sheet must be referenced here.
     */
    private void formatSecondSheet(XSSFSheet sheet) {
        secondSheetTopLabel(sheet);
        secondSheetTopLeftTable(sheet);
        secondSheetTopRightTable(sheet);
        secondSheetBottomLeftTable(sheet);
        secondSheetBottomRightTable(sheet);
    }

    /**
     * Formats the third sheet, calling a variety of different methods,
     * formatting different areas of the sheet
     *
     * @param sheet - the third sheet must be referenced here.
     */
    private void formatThirdSheet(XSSFSheet sheet) {
        thirdSheetTopLabel(sheet);
        thirdSheetRowStatistics(sheet);
    }

    /**
     * Formats the fourth sheet, calling a variety of different methods,
     * formatting different areas of the sheet
     *
     * @param sheet - the fourth sheet must be referenced here.
     */
    private void formatFourthSheet(XSSFSheet sheet) {
        fourthSheetTopLabel(sheet);
        fourthSheetRowStatistics(sheet);
    }

    //--------------FIRST SHEET FORMATTERS-----------------//
    private void firstSheetSelloutLabel(XSSFSheet sheet) {
        String labelAddress = "A1:BE1";
        formatLabel(Constants.SELL_OUT, labelAddress, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetStockLabel(XSSFSheet sheet) {
        String labelAddress = "BG1:BJ1";
        formatLabel(Constants.STOCK, labelAddress, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetTotalLabel(XSSFSheet sheet) {
        String labelAddress = "BL1:BO1";
        formatLabel(Constants.TOTAL, labelAddress, sheet, 0xDAEEF3, (short) 0);
    }

    private void firstSheetSelloutNamcoLabel(XSSFSheet sheet) {
        String labelAddress = "A" + Constants.PLATFORMS_FIRSTROW + ":A" + Constants.PLATFORMS_TABLE_LASTROW;
        formatLabel(Constants.NAMCO, labelAddress, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetStockNamcoLabel(XSSFSheet sheet) {
        String labelAddress = "BG" + Constants.PLATFORMS_FIRSTROW + ":BG" + Constants.PLATFORMS_TABLE_LASTROW;
        formatLabel(Constants.NAMCO, labelAddress, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetTotalNamcoLabel(XSSFSheet sheet) {
        String labelAddress = "BL" + Constants.PLATFORMS_FIRSTROW + ":BL" + Constants.PLATFORMS_TABLE_LASTROW;
        formatLabel(Constants.NAMCO, labelAddress, sheet, 0xDA9694, (short) 90);
    }

    private void firstSheetTableSellout(XSSFSheet sheet) {
        CellRangeAddress tableAddress = CellRangeAddress.valueOf("B" + Constants.PLATFORMS_TABLE_WEEK_ROW + ":BE" + Constants.PLATFORMS_TABLE_LASTROW);
        createCells(tableAddress, sheet);
        XSSFCellStyle tableCellStyle = (XSSFCellStyle) wb.createCellStyle();
        applyBorderStyle(tableCellStyle, BorderStyle.THIN);
        applyStyleToCells(tableAddress, sheet, tableCellStyle);
        XSSFCellStyle totalBarCellStyle = (XSSFCellStyle) wb.createCellStyle();
        totalBarCellStyle.cloneStyleFrom(tableCellStyle);
        Font tableFont = wb.createFont();
        tableFont.setBold(true);
        tableFont.setFontHeightInPoints((short) 12);
        totalBarCellStyle.setFont(tableFont);
        applyStyleToCells(CellRangeAddress.valueOf("B" + Constants.PLATFORMS_TABLE_LASTROW + ":BC" + Constants.PLATFORMS_TABLE_LASTROW), sheet, totalBarCellStyle);
        XSSFCellStyle weekBarStyle = (XSSFCellStyle) wb.createCellStyle();
        weekBarStyle.cloneStyleFrom(tableCellStyle);
        align(weekBarStyle, HorizontalAlignment.RIGHT);
        applyStyleToCells(CellRangeAddress.valueOf("C" + Constants.PLATFORMS_TABLE_WEEK_ROW + ":BB" + Constants.PLATFORMS_TABLE_WEEK_ROW), sheet, weekBarStyle);
        //Sets the platforms labels to the left and right
        int rowIter = Constants.PLATFORMS_FIRSTROW;
        for (Platforms platform : Platforms.values()) {
            CellReference leftCellRef = new CellReference("B" + rowIter);
            CellReference rightCellRef = new CellReference("BC" + rowIter);
            Cell leftCell = sheet.getRow(leftCellRef.getRow()).getCell(leftCellRef.getCol());
            leftCell.setCellValue(platform.toString());
            Cell rightCell = sheet.getRow(rightCellRef.getRow()).getCell(rightCellRef.getCol());
            rightCell.setCellValue(platform.toString());
            rowIter++;
        }
        //Sets the total labels to the left and right
        CellReference totalLeftCellRef = new CellReference("B" + Constants.PLATFORMS_TABLE_LASTROW);
        CellReference totalRightCellRef = new CellReference("BC" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalLeftCell = sheet.getRow(totalLeftCellRef.getRow()).getCell(totalLeftCellRef.getCol());
        Cell totalRightCell = sheet.getRow(totalRightCellRef.getRow()).getCell(totalRightCellRef.getCol());
        totalLeftCell.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        totalRightCell.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        //Sets the weekly sum formulas
        for (int column = CellReference.convertColStringToIndex("C"); column <= CellReference.convertColStringToIndex("BB"); column++) {
            String letter = CellReference.convertNumToColString(column);
            CellReference totalCellRef = new CellReference(letter + Constants.PLATFORMS_TABLE_LASTROW);
            Cell totalCell = sheet.getRow(totalCellRef.getRow()).getCell(totalCellRef.getCol());
            totalCell.setCellType(CellType.FORMULA);
            totalCell.setCellFormula("SUM(" + letter + Constants.PLATFORMS_FIRSTROW + ":" + letter + Constants.PLATFORMS_LASTROW + ")");
        }
        XSSFCellStyle totalPcsStyle = (XSSFCellStyle) wb.createCellStyle();
        totalPcsStyle.cloneStyleFrom(tableCellStyle);
        //Same as before
        totalBarCellStyle.setFont(tableFont);
        applyStyleToCells(CellRangeAddress.valueOf("BD" + Constants.PLATFORMS_TABLE_WEEK_ROW), sheet, totalPcsStyle);
        CellReference totalPcs = new CellReference("BD" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Cell totalPcsLabelCell = sheet.getRow(totalPcs.getRow()).getCell(totalPcs.getCol());
        totalPcsLabelCell.setCellValue(Constants.TOTAL_PCS);
        //Sets color of the total pcs column
        XSSFCellStyle totalPcsColumnStyle = (XSSFCellStyle) wb.createCellStyle();
        totalPcsColumnStyle.cloneStyleFrom(tableCellStyle);
        setBackgroundColor(totalPcsColumnStyle, 0x92D050);
        applyStyleToCells(CellRangeAddress.valueOf("BD" + Constants.PLATFORMS_FIRSTROW + ":BD" + Constants.PLATFORMS_TABLE_LASTROW), sheet, totalPcsColumnStyle);
        //From platforms firstrow to the last row of the table, so i can sum all weeks
        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_TABLE_LASTROW; i++) {
            CellReference totalPcsCellRef = new CellReference("BD" + i);
            Cell totalPcsCell = sheet.getRow(totalPcsCellRef.getRow()).getCell(totalPcsCellRef.getCol());
            totalPcsCell.setCellType(CellType.FORMULA);
            totalPcsCell.setCellFormula("SUM(C" + i + ":BB" + i + ")");
        }
        //Percentage column formatting
        XSSFCellStyle percentageColumnStyle = (XSSFCellStyle) wb.createCellStyle();
        percentageColumnStyle.cloneStyleFrom(tableCellStyle);
        percentageColumnStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        applyStyleToCells(CellRangeAddress.valueOf("BE" + Constants.PLATFORMS_FIRSTROW + ":BE" + Constants.PLATFORMS_TABLE_LASTROW), sheet, percentageColumnStyle);
        //Percentage column formulas
        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_LASTROW; i++) {
            CellReference percentageCellsRef = new CellReference("BE" + i);
            Cell percentageCell = sheet.getRow(percentageCellsRef.getRow()).getCell(percentageCellsRef.getCol());
            percentageCell.setCellType(CellType.FORMULA);
            percentageCell.setCellFormula("BD" + i + "/BD" + Constants.PLATFORMS_TABLE_LASTROW);
        }
        //Total percentage formula (adds up to 100%)
        CellReference totalPercentageCellRef = new CellReference("BE" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalPercentageCell = sheet.getRow(totalPercentageCellRef.getRow()).getCell(totalPercentageCellRef.getCol());
        totalPercentageCell.setCellType(CellType.FORMULA);
        totalPercentageCell.setCellFormula("SUM(BE" + Constants.PLATFORMS_FIRSTROW + ":BE" + Constants.PLATFORMS_LASTROW + ")");
    }

    private void firstSheetTableStockFormat(XSSFSheet sheet) {
        CellRangeAddress tableAddress = CellRangeAddress.valueOf("BH" + Constants.PLATFORMS_FIRSTROW + ":" + "BJ" + Constants.PLATFORMS_TABLE_LASTROW);
        createCells(tableAddress, sheet);
        XSSFCellStyle tableCellStyle = (XSSFCellStyle) wb.createCellStyle();
        applyBorderStyle(tableCellStyle, BorderStyle.THIN);
        applyStyleToCells(tableAddress, sheet, tableCellStyle);
        //Adds platform labels to table stock
        int rowIter = Constants.PLATFORMS_FIRSTROW;
        for (Platforms platform : Platforms.values()) {
            CellReference platformLabelCellRef = new CellReference("BH" + rowIter);
            CellReference platformDaysInStockCellRef = new CellReference("BJ" + rowIter);
            Cell platformLabelCell = sheet.getRow(platformLabelCellRef.getRow()).getCell(platformLabelCellRef.getCol());
            Cell platformDaysInStockCell = sheet.getRow(platformDaysInStockCellRef.getRow()).getCell(platformDaysInStockCellRef.getCol());
            platformLabelCell.setCellValue(platform.toString());
            platformDaysInStockCell.setCellType(CellType.FORMULA);
            platformDaysInStockCell.setCellFormula("IF(OR(BI" + rowIter + "=\"No data\",BI" + rowIter + "=\"\"),BI" + rowIter + "&\"\",IFERROR(BI" + rowIter + 
                    "/BD" + rowIter + "*7*COUNT(C" + rowIter + ":BB" + rowIter + "),\"\"))");
            rowIter++;
        }
        //Sets the Total label
        CellReference totalLabelCellRef = new CellReference("BH" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalLabelCell = sheet.getRow(totalLabelCellRef.getRow()).getCell(totalLabelCellRef.getCol());
        totalLabelCell.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        //Sets the Total formula
        CellReference totalFormulaCellRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalFormulaCell = sheet.getRow(totalFormulaCellRef.getRow()).getCell(totalFormulaCellRef.getCol());
        totalFormulaCell.setCellType(CellType.FORMULA);
        totalFormulaCell.setCellFormula("SUM(BI" + Constants.PLATFORMS_FIRSTROW + ":BI" + Constants.PLATFORMS_LASTROW + ")");
        //Sets the days in stock label
        CellReference daysInStockLabelCellRef = new CellReference("BJ" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Cell daysInStockLabelCell = sheet.getRow(daysInStockLabelCellRef.getRow()).getCell(daysInStockLabelCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        daysInStockLabelCell.setCellValue(Constants.DAYS_IN_STOCK);
        //Sets the days in stock number format
        XSSFCellStyle daysInStockNumberFormat = (XSSFCellStyle) wb.createCellStyle();
        daysInStockNumberFormat.cloneStyleFrom(tableCellStyle);
        daysInStockNumberFormat.setDataFormat(wb.createDataFormat().getFormat("0"));
        applyStyleToCells(CellRangeAddress.valueOf("BJ" + Constants.PLATFORMS_FIRSTROW + ":BJ" + Constants.PLATFORMS_TABLE_LASTROW), sheet, daysInStockNumberFormat);

        CellReference totalDaysInStockRef = new CellReference("BJ" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalDaysInStockCell = sheet.getRow(totalDaysInStockRef.getRow()).getCell(totalDaysInStockRef.getCol());
        totalDaysInStockCell.setCellType(CellType.FORMULA);
        totalDaysInStockCell.setCellFormula("IFERROR(BI" + Constants.PLATFORMS_TABLE_LASTROW + "/BD" + Constants.PLATFORMS_TABLE_LASTROW
                + "*7*" + "COUNTIF(C" + Constants.PLATFORMS_TABLE_WEEK_ROW + ":BB" + Constants.PLATFORMS_TABLE_WEEK_ROW + ",\"<>\"&\"\"),0)");
    }

    private void firstSheetTableTotalFormat(XSSFSheet sheet) {
        CellRangeAddress tableAddress = CellRangeAddress.valueOf("BM" + Constants.PLATFORMS_FIRSTROW + ":" + "BO" + Constants.PLATFORMS_TABLE_LASTROW);
        createCells(tableAddress, sheet);
        XSSFCellStyle tableCellStyle = (XSSFCellStyle) wb.createCellStyle();
        applyBorderStyle(tableCellStyle, BorderStyle.THIN);
        applyStyleToCells(tableAddress, sheet, tableCellStyle);

        //Sets the platform labels to the table
        int rowIter = Constants.PLATFORMS_FIRSTROW;
        for (Platforms platform : Platforms.values()) {
            CellReference platformCellRef = new CellReference("BM" + rowIter);
            Cell platformCell = sheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
            platformCell.setCellValue(platform.toString());
            rowIter++;
        }
        //Sets the Total label
        CellReference totalLabelCellRef = new CellReference("BM" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalLabelCell = sheet.getRow(totalLabelCellRef.getRow()).getCell(totalLabelCellRef.getCol());
        totalLabelCell.setCellValue(Constants.PLATFORMS_TABLE_TOTAL);
        //Sets the Sales label
        CellReference salesLabelRef = new CellReference("BN" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Cell salesLabelCell = sheet.getRow(salesLabelRef.getRow()).getCell(salesLabelRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        salesLabelCell.setCellValue(Constants.SALES);
        //Sets the percentage column style
        XSSFCellStyle percentageStyle = (XSSFCellStyle) wb.createCellStyle();
        percentageStyle.cloneStyleFrom(tableCellStyle);
        percentageStyle.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        //Applies percentage style to the sales column
        applyStyleToCells(CellRangeAddress.valueOf("BN" + Constants.PLATFORMS_FIRSTROW + ":BN" + Constants.PLATFORMS_TABLE_LASTROW), sheet, percentageStyle);

        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_TABLE_LASTROW; i++) {
            CellReference salesCellRef = new CellReference("BN" + i);
            Cell salesCell = sheet.getRow(salesCellRef.getRow()).getCell(salesCellRef.getCol());
            salesCell.setCellType(CellType.FORMULA);
            salesCell.setCellFormula("BE" + i);
        }
        //Sets the total label
        CellReference stockLabelRef = new CellReference("BO" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Cell stockLabel = sheet.getRow(stockLabelRef.getRow()).getCell(stockLabelRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        stockLabel.setCellValue(Constants.STOCK);
        //Applies percentage style to the stock column
        applyStyleToCells(CellRangeAddress.valueOf("BO" + Constants.PLATFORMS_FIRSTROW + ":BO" + Constants.PLATFORMS_TABLE_LASTROW), sheet, percentageStyle);
        for (int i = Constants.PLATFORMS_FIRSTROW; i <= Constants.PLATFORMS_LASTROW; i++) {
            CellReference stockCellRef = new CellReference("BO" + i);
            Cell stockCell = sheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol());
            stockCell.setCellType(CellType.FORMULA);
            stockCell.setCellFormula("BI" + i + "/BI" + Constants.PLATFORMS_TABLE_LASTROW);
        }
        //Total stock percentage formula. Adds up to 100%
        CellReference totalStockCellRef = new CellReference("BO" + Constants.PLATFORMS_TABLE_LASTROW);
        Cell totalStockCell = sheet.getRow(totalStockCellRef.getRow()).getCell(totalStockCellRef.getCol());
        totalStockCell.setCellType(CellType.FORMULA);
        totalStockCell.setCellFormula("SUM(BO" + Constants.PLATFORMS_FIRSTROW + ":BO" + Constants.PLATFORMS_LASTROW + ")");
    }

    //--------------END OF FIRST SHEET FORMATTERS-----------------//
    //--------------SECOND SHEET FORMATTERS-----------------//
    private void secondSheetTopLabel(XSSFSheet sheet) {
        String cellRange = "A1:R1";
        formatLabel(Constants.SHEET_2_LABEL, cellRange, sheet, 0xDAEEF3, (short) 0);
    }

    private void secondSheetTopLeftTable(XSSFSheet sheet) {
        String labelAddress = "C5:H11";
        secondSheetFormatTable(Constants.TOP_LEFT_LABEL, Constants.SHOP, labelAddress, sheet);
    }

    private void secondSheetTopRightTable(XSSFSheet sheet) {
        String labelAddress = "K5:P11";
        secondSheetFormatTable(Constants.TOP_RIGHT_LABEL, Constants.SHOP, labelAddress, sheet);
    }

    private void secondSheetBottomLeftTable(XSSFSheet sheet) {
        String labelAddress = "C14:H20";
        secondSheetFormatTable(Constants.BOTTOM_LEFT_LABEL, Constants.GAME, labelAddress, sheet);
    }

    private void secondSheetBottomRightTable(XSSFSheet sheet) {
        String labelAddress = "K14:P20";
        secondSheetFormatTable(Constants.BOTTOM_RIGHT_LABEL, Constants.GAME, labelAddress, sheet);
    }

    private void secondSheetFormatTable(String tableLabel, String secondRowLabel, String cellRange, XSSFSheet sheet) {
        CellRangeAddress tableAddress = CellRangeAddress.valueOf(cellRange);
        createCells(tableAddress, sheet);

        XSSFCellStyle tableStyle = (XSSFCellStyle) wb.createCellStyle();
        applyBorderStyle(tableStyle, BorderStyle.THIN);
        applyStyleToCells(tableAddress, sheet, tableStyle);

        formatLabel(tableLabel, new CellRangeAddress(tableAddress.getFirstRow(), tableAddress.getFirstRow(), tableAddress.getFirstColumn(), tableAddress.getLastColumn())
                .formatAsString(), sheet, 0xFFA500, (short) 0);

        XSSFCellStyle firstRowLabelStyle = (XSSFCellStyle) wb.createCellStyle();
        firstRowLabelStyle.cloneStyleFrom(tableStyle);

        Font labelFont = wb.createFont();
        labelFont.setBold(true);
        firstRowLabelStyle.setFont(labelFont);

        for (int i = tableAddress.getFirstRow() + 1; i <= tableAddress.getLastRow(); i++) {
            CellRangeAddress secondRowLabelAddress = new CellRangeAddress(i, i, tableAddress.getFirstColumn(), tableAddress.getLastColumn() - 1);
            mergeCells(secondRowLabelAddress, sheet);
            applyStyleToCells(secondRowLabelAddress, sheet, tableStyle);
        }

        XSSFCellStyle secondRowStyle = (XSSFCellStyle) wb.createCellStyle();
        secondRowStyle.cloneStyleFrom(tableStyle);
        secondRowStyle.setFont(labelFont);
        align(secondRowStyle, HorizontalAlignment.CENTER);

        Row secondRow = sheet.getRow(tableAddress.getFirstRow() + 1);
        Cell secondRowLeftCell = secondRow.getCell(tableAddress.getFirstColumn());
        secondRowLeftCell.setCellValue(secondRowLabel);
        applyStyleToCells(new CellRangeAddress(tableAddress.getFirstRow() + 1, tableAddress.getFirstRow() + 1, tableAddress.getFirstColumn(), tableAddress.getLastColumn() - 1), sheet, secondRowStyle);

        Cell secondRowRightCell = secondRow.getCell(tableAddress.getLastColumn());
        secondRowRightCell.setCellValue(Constants.SALES);
        applyStyleToCells(new CellRangeAddress(tableAddress.getFirstRow() + 1, tableAddress.getFirstRow() + 1, tableAddress.getLastColumn(), tableAddress.getLastColumn()), sheet, secondRowStyle);
    }

    //--------------END OF SECOND SHEET FORMATTERS-----------------//
    //--------------THIRD SHEET FORMATTERS----------------------//
    private void thirdSheetTopLabel(XSSFSheet sheet) {
        String labelAddress = "A1:R1";
        formatLabel(Constants.SHEET_3_LABEL, labelAddress, sheet, 0xDAEEF3, (short) 0);
    }

    private void thirdSheetRowStatistics(XSSFSheet sheet) {
        final int row = 2; //0 based
        final int column = 3; //0 based

        CellRangeAddress region = new CellRangeAddress(row, row, column, column + Platforms.values().length);
        createCells(region, sheet);
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        align(style, HorizontalAlignment.CENTER);
        applyStyleToCells(region, sheet, style);
        int iterator = column;
        for (Platforms platform : Platforms.values()) {
            Cell cell = sheet.getRow(row).getCell(iterator, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            cell.setCellValue(platform.toString());
            iterator++;
        }
        Row thirdRow = sheet.getRow(row);
        Cell totalCell = thirdRow.getCell(column + Platforms.values().length);
        totalCell.setCellValue(Constants.TOTAL);
    }

    //--------------END OF THIRD SHEET FORMATTERS----------------//
    //--------------FOURTH SHEET FORMATTERS---------------//
    private void fourthSheetTopLabel(XSSFSheet sheet) {
        String labelAddress = "A1:R1";
        formatLabel(Constants.SHEET_4_LABEL, labelAddress, sheet, 0xDAEEF3, (short) 0);
    }

    private void fourthSheetRowStatistics(XSSFSheet sheet) {
        CellRangeAddress platformCellAddress = CellRangeAddress.valueOf("A3");
        CellUtil.getRow(platformCellAddress.getFirstRow(), sheet);

        Cell platformCell = sheet.getRow(platformCellAddress.getFirstRow()).getCell(platformCellAddress.getFirstColumn(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        platformCell.setCellValue(Constants.PLATFORM);

        CellRangeAddress gameCellsAddress = CellRangeAddress.valueOf("B3:D3");
        createCells(gameCellsAddress, sheet);
        mergeCells(gameCellsAddress, sheet);
        setMergedCellsValue(gameCellsAddress, sheet, Constants.GAME);

        CellRangeAddress salesCellAddress = CellRangeAddress.valueOf("E3");
        Cell salesCell = sheet.getRow(salesCellAddress.getFirstRow()).getCell(salesCellAddress.getFirstColumn(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        salesCell.setCellValue(Constants.SALES);

    }

    //--------------END OF FOURTH SHEET FORMATTERS-------------//
    //--------------TOOLS-----------------
    private void formatLabel(String labelText, String cellRange, Sheet sheet, int backgroundColor, short rotationDegrees) {
        CellRangeAddress region = CellRangeAddress.valueOf(cellRange);
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        applyBorderStyle(style, BorderStyle.THIN);
        align(style, HorizontalAlignment.CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        setRotation(style, rotationDegrees);
        setBackgroundColor(style, backgroundColor);
        createCells(region, sheet);
        applyStyleToCells(region, sheet, style);
        mergeCells(region, sheet);
        setMergedCellsValue(region, sheet, labelText);
    }

    /**
     * Creates a border around the selected cells
     *
     * @param style Style to apply the border to.
     * @param borderStyle Type of border.
     */
    private void applyBorderStyle(XSSFCellStyle style, BorderStyle borderStyle) {
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
    private void applyStyleToCells(CellRangeAddress region, Sheet sheet, XSSFCellStyle style) {
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
