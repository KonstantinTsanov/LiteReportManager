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
import java.util.logging.Level;
import java.util.logging.Logger;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.java.Log;
import net.thecir.constants.Constants;
import net.thecir.enums.ExcelWorkbookType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
    protected XSSFSheet reviewSheet;
    protected XSSFSheet topFiveSheet;
    protected XSSFSheet salesByPlatformSheet;
    protected XSSFSheet salesByGameSheet;

    //Indicated whether the user is adding or removing
    protected boolean undo;

    /**
     * Shop, platform, game, stock/sales.
     */
    protected HashMap<String, HashMap<String, HashMap<String, StockSales>>> newData = new HashMap<>();

    public ReportManager(File inputFilePath, File outputFilePath, boolean undo) {
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
                log.log(Level.SEVERE, null, ex);
            } catch (InvalidFormatException ex) {
                log.log(Level.SEVERE, null, ex);
            }
        }

    }

    protected void writeToSheet() {
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

    private void writeWeeklyReport() {
        for (int column = Constants.SELLOUTTABLEFIRSTCOLUMN; column <= Constants.SELLOUTTABLELASTCOLUMN; column++) {

        }
    }

    private void undoWeeklyReport() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private HashMap<String, StockSales> getStockSalesPerPlatform() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private int getColumnToDelete(HashMap<String, StockSales> newData, int weekNo) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void clearTopFiveStatistics() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    protected int platformWeeksInStock(int row) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private int totalNumberOfWeeksPresent() {
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
