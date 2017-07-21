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
package net.thecir.reportmanagers;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.java.Log;
import net.thecir.constants.Constants;
import net.thecir.enums.ExcelWorkbookType;
import net.thecir.enums.Platforms;
import net.thecir.exceptions.OutputFileIsFullException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public abstract class ReportManager {

    /**
     * Used internally as data structure that holds stock and sales for each
     * game.
     */
    protected class StockSales {

        @Getter
        @Setter
        int Stock;

        @Getter
        @Setter
        int Sales;
    }
    //Input and output sheets
    private Workbook inputWorkbook;
    private XSSFWorkbook outputWorkbook;
    //xls or xlsx
    private ExcelWorkbookType inputWorkbookType;

    //Input worksheet
    protected Sheet inputDataSheet;

    //Output worksheets
    protected XSSFSheet weeklyReportSheet;
    protected XSSFSheet topFiveSheet;
    protected XSSFSheet salesByPlatformSheet;
    protected XSSFSheet salesByGameSheet;

    //Indicated whether the user is adding or removing
    protected boolean undo;

    ResourceBundle rb;
    /**
     * Shop, platform, game, stock/sales.
     */
    protected HashMap<String, HashMap<String, HashMap<String, StockSales>>> newData = new HashMap<>();

    public ReportManager(File inputFilePath, File outputFilePath, boolean undo) {
        this.undo = undo;
        rb = ResourceBundle.getBundle("CoreBundle");
        try {
            outputWorkbook = new XSSFWorkbook(outputFilePath);
        } catch (IOException ex) {
            //TODO
            log.log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            log.log(Level.SEVERE, null, ex);
        }
        if (inputWorkbookType == ExcelWorkbookType.XLS) {
            try (InputStream is = new FileInputStream(inputFilePath)) {
                inputWorkbook = new HSSFWorkbook(is);
            } catch (FileNotFoundException ex) {
                //TODO
                log.log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                log.log(Level.SEVERE, null, ex);
            }
        } else {
            try {
                outputWorkbook = new XSSFWorkbook(inputFilePath);
            } catch (IOException ex) {
                //TODO
                log.log(Level.SEVERE, null, ex);
            } catch (InvalidFormatException ex) {
                log.log(Level.SEVERE, null, ex);
            }
        }

    }

    protected void writeToSheet() throws OutputFileIsFullException {
        if (!undo) {
            writeWeeklyReport();
        } else {
            undoWeeklyReport();
            clearTopFiveStatistics();
        }
        writeOverallSalesByPlatform();
        writeOverallSalesByGame();
        writeTopFiveStatistics();
    }

    private void writeWeeklyReport() throws OutputFileIsFullException {
        int weekNo = getWeekNumber();
        for (int column = Constants.SELLOUT_TABLE_FIRST_COLUMN; column <= Constants.SELLOUT_TABLE_LAST_COLUMN; column++) {
            CellReference weekCellRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW, column);
            if (weeklyReportSheet.getRow(weekCellRef.getRow()).getCell(weekCellRef.getCol()).getCellTypeEnum() != CellType.BLANK) {
                if (column == Constants.SELLOUT_TABLE_LAST_COLUMN) {
                    throw new OutputFileIsFullException(rb.getString("OutputFileIsFullExceptionMessage"));
                }
                continue;
            }
            HashMap<String, StockSales> stockAndSalesByPlatform = getStockSalesByPlatform();
            CellReference latestWeekStockCellRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_WEEK_ROW);
            //Set the current report's week
            weeklyReportSheet.getRow(latestWeekStockCellRef.getRow()).getCell(latestWeekStockCellRef.getCol()).setCellValue("Stock w" + weekNo);
            weeklyReportSheet.getRow(weekCellRef.getRow()).getCell(weekCellRef.getCol()).setCellValue("w" + weekNo);
            for (int row = Constants.PLATFORMS_LASTROW; row < Platforms.values().length; row++) {
                CellReference currentOutputAbbreviationRef = new CellReference("B" + row);
                String currentOutputAbbreviation = weeklyReportSheet.getRow(currentOutputAbbreviationRef.getRow())
                        .getCell(currentOutputAbbreviationRef.getCol()).getStringCellValue();
                if (stockAndSalesByPlatform.get(currentOutputAbbreviation).Sales != Integer.MIN_VALUE) {
                    CellReference currentCellRef = new CellReference(row, column);
                    weeklyReportSheet.getRow(currentCellRef.getRow()).getCell(currentCellRef.getCol())
                            .setCellValue(stockAndSalesByPlatform.get(currentOutputAbbreviation).Sales);
                }
                CellReference stockCellRef = new CellReference("BI" + row);
                weeklyReportSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellType(CellType.BLANK);
                //If records about this platform exist in the latest report proceed.
                if (stockAndSalesByPlatform.get(currentOutputAbbreviation).Stock != Integer.MIN_VALUE) {
                    weeklyReportSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol())
                            .setCellValue(stockAndSalesByPlatform.get(currentOutputAbbreviation).Stock);
                }
            }
            break;
        }
    }

    private void undoWeeklyReport() {
        int weekNo = getWeekNumber();
        HashMap<String, StockSales> stockAndSalesByPlatform = getStockSalesByPlatform();
        int columnToRemove = getColumnToRemove(stockAndSalesByPlatform, weekNo);
        CellReference cellOfWeekToRemoveRef = new CellReference(3, columnToRemove);
        weeklyReportSheet.getRow(cellOfWeekToRemoveRef.getRow()).getCell(cellOfWeekToRemoveRef.getCol()).setCellType(CellType.BLANK);
        for (int row = Constants.PLATFORMS_FIRSTROW; row < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; row++) {
            CellReference cellToRemoveRef = new CellReference(row, columnToRemove);
            weeklyReportSheet.getRow(cellToRemoveRef.getRow()).getCell(cellToRemoveRef.getCol()).setCellType(CellType.BLANK);
        }
        //Checking if the lastly added week is the week to be removed.
        CellReference stockWeekNumberCellRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Pattern pattern = Pattern.compile("w" + weekNo);
        Matcher m = pattern.matcher(weeklyReportSheet.getRow(stockWeekNumberCellRef.getRow()).getCell(stockWeekNumberCellRef.getCol()).getStringCellValue());
        if (m.matches()) {
            for (int row = Constants.PLATFORMS_FIRSTROW; row < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; row++) {
                CellReference latestWeekStockCellRef = new CellReference("BI" + row);
                CellReference currentRowPlatformCellRef = new CellReference("B" + row);
                if (weeklyReportSheet.getRow(latestWeekStockCellRef.getRow()).getCell(latestWeekStockCellRef.getCol()).getCellTypeEnum() == CellType.BLANK) {
                    if (stockAndSalesByPlatform.get(weeklyReportSheet.getRow(currentRowPlatformCellRef
                            .getRow()).getCell(currentRowPlatformCellRef.getCol()).getStringCellValue()).Stock == 0) {
                        continue;
                    }
                } else {
                    int stockParser = Integer.parseInt(weeklyReportSheet.getRow(latestWeekStockCellRef.getRow()).getCell(latestWeekStockCellRef.getCol()).getStringCellValue());
                    if (stockParser == stockAndSalesByPlatform.get(weeklyReportSheet.getRow(currentRowPlatformCellRef
                            .getRow()).getCell(currentRowPlatformCellRef.getCol()).getStringCellValue()).Stock) {
                        continue;
                    }
                    break;
                }
                for (int rowToDeleteOn = Constants.PLATFORMS_FIRSTROW; rowToDeleteOn < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; rowToDeleteOn++) {
                    CellReference stockCellToBeRemoved = new CellReference("BI" + rowToDeleteOn);
                    weeklyReportSheet.getRow(stockCellToBeRemoved.getRow()).getCell(stockCellToBeRemoved.getCol()).setCellValue(Constants.NO_DATA);
                }
            }
        }
    }

    private HashMap<String, StockSales> getStockSalesByPlatform() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private int getColumnToRemove(HashMap<String, StockSales> newData, int weekNo) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void clearTopFiveStatistics() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private boolean outputFileSignature() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void writeTopFiveStatistics() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void topFiveShopsBySalesOverall() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void topFiveGamesBySalesOverall() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void topFiveShopsBySalesLatestWeek() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void topFiveGamesBySalesLatestWeek() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void writeOverallSalesByPlatform() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void overallSalesByPlatformExistingRecords() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void overallSalesByPlatformFreshRecords() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private HashMap<String, HashMap<String, Integer>> getCurrentOverallSalesByPlatform() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void writeOverallSalesByGame() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void salesByGameExistingRecords() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void salesByGameFreshRecords() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    protected void formatDataHashMap() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    //TODO abstract methods!

    /**
     * Get week number from the input file.
     *
     * @return week number.
     */
    protected abstract int getWeekNumber();
}
