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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.java.Log;
import net.thecir.constants.Constants;
import net.thecir.enums.ExcelWorkbookType;
import net.thecir.enums.Platforms;
import net.thecir.exceptions.InputFileContainsNoValidDateException;
import net.thecir.exceptions.OutputFileIsFullException;
import net.thecir.exceptions.OutputFileNoRecordsFoundException;
import net.thecir.exceptions.OutputFileNotCorrectException;
import net.thecir.exceptions.InputFileNotMatchingSelectedFileException;
import net.thecir.exceptions.OutputFileIOException;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
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
    private Workbook outputWorkbook;

    private File outputWorkbookFile;

    //Input worksheet
    protected Sheet inputDataSheet;

    //Output worksheets
    protected Sheet weeklyReportSheet;
    protected Sheet topFiveSheet;
    protected Sheet salesByPlatformSheet;
    protected Sheet salesByGameSheet;

    //Evaluator is needed to evaluate the cells before getting the value, otherwise we get incorrect results.
    protected FormulaEvaluator evaluator;
    //Indicated whether the user is adding or removing
    protected boolean undo;

    ResourceBundle rb;
    /**
     * Shop, platform, game, stock/sales.
     */
    protected HashMap<String, HashMap<String, HashMap<String, StockSales>>> newData;

    public ReportManager(File inputWorkbookFile, File outputWorkbookFile, boolean undo) {
        this.outputWorkbookFile = outputWorkbookFile;
        this.undo = undo;
        newData = new HashMap<>();
        rb = ResourceBundle.getBundle("LanguageBundles/Bundle");
        try {
            //TODO Input stream resource
            outputWorkbook = WorkbookFactory.create(new FileInputStream(outputWorkbookFile));
        } catch (IOException ex) {
            //TODO
            log.log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            log.log(Level.SEVERE, null, ex);
        }

        try (InputStream is = new BufferedInputStream(new FileInputStream(inputWorkbookFile))) {
            if (POIFSFileSystem.hasPOIFSHeader(is)) {
                inputWorkbook = new HSSFWorkbook(is);
            } else if (DocumentFactoryHelper.hasOOXMLHeader(is)) {
                inputWorkbook = new XSSFWorkbook(is);
            }
        } catch (IOException ex) {
            Logger.getLogger(ReportManager.class.getName()).log(Level.SEVERE, null, ex);
        }
        inputDataSheet = inputWorkbook.getSheetAt(0);
        weeklyReportSheet = outputWorkbook.getSheetAt(0);
        topFiveSheet = outputWorkbook.getSheetAt(1);
        salesByPlatformSheet = outputWorkbook.getSheetAt(2);
        salesByGameSheet = outputWorkbook.getSheetAt(3);
        evaluator = outputWorkbook.getCreationHelper().createFormulaEvaluator();
    }

    public void generateReport() throws OutputFileIsFullException,
            OutputFileNoRecordsFoundException, InputFileNotMatchingSelectedFileException,
            OutputFileNotCorrectException, OutputFileIOException, InputFileContainsNoValidDateException {
        if (!isOutputFileCorrect()) {
            throw new OutputFileNotCorrectException("The destionation file isn't correct. Select another file or create new.");
        }
        if (!isInputFileCorrect()) {
            throw new InputFileNotMatchingSelectedFileException("The source file isn't from the selected retailer.");
        }
        formatDataHashMap();
        readInputData();
        writeToSheet();
        try (FileOutputStream fileOut = new FileOutputStream(outputWorkbookFile)) {
            XSSFFormulaEvaluator.evaluateAllFormulaCells(outputWorkbook);
            outputWorkbook.write(fileOut);
        } catch (FileNotFoundException ex) {
            log.log(Level.SEVERE, "The file to save the workbook in was not found.", ex);
            throw new OutputFileIOException("The file to save the workbook in was not found.");
        } catch (IOException ex) {
            log.log(Level.SEVERE, "There's an IO problem with the output file.", ex);
            throw new OutputFileIOException("There's an IO problem with the output file.");
        }
    }

    protected void writeToSheet() throws OutputFileIsFullException, OutputFileNoRecordsFoundException, InputFileContainsNoValidDateException {
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

    private void writeWeeklyReport() throws OutputFileIsFullException, InputFileContainsNoValidDateException {
        int weekNo = getWeekNumber();
        for (int column = Constants.SELLOUT_TABLE_FIRST_COLUMN; column <= Constants.SELLOUT_TABLE_LAST_COLUMN; column++) {
            CellReference weekCellRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW - 1, column - 1);
            if (!"".equals(weeklyReportSheet.getRow(weekCellRef.getRow()).getCell(weekCellRef.getCol()).getStringCellValue())) {
                if (column == Constants.SELLOUT_TABLE_LAST_COLUMN) {
                    throw new OutputFileIsFullException(rb.getString("OutputFileIsFullExceptionMessage"));
                }
                continue;
            }
            HashMap<String, StockSales> stockAndSalesByPlatform = getStockSalesByPlatform();
            CellReference latestWeekStockCellRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_WEEK_ROW);
            //Set the current report's week
            CellUtil.getRow(latestWeekStockCellRef.getRow(), weeklyReportSheet)
                    .getCell(latestWeekStockCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue("Stock w" + weekNo);
            weeklyReportSheet.getRow(weekCellRef.getRow()).getCell(weekCellRef.getCol()).setCellValue("w" + weekNo);
            for (int row = Constants.PLATFORM_HEADER_FIRST_ROW; row <= Constants.PLATFORM_HEADER_LAST_ROW; row++) {
                CellReference currentOutputAbbreviationRef = new CellReference("B" + row);
                String currentOutputAbbreviation = weeklyReportSheet.getRow(currentOutputAbbreviationRef.getRow())
                        .getCell(currentOutputAbbreviationRef.getCol()).getStringCellValue();
                if (stockAndSalesByPlatform.get(currentOutputAbbreviation).Sales != Integer.MIN_VALUE) {
                    CellReference currentCellRef = new CellReference(row - 1, column - 1);
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

    private void undoWeeklyReport() throws OutputFileNoRecordsFoundException, InputFileContainsNoValidDateException {
        int weekNo = getWeekNumber();
        HashMap<String, StockSales> stockAndSalesByPlatform = getStockSalesByPlatform();
        int columnToRemove = findWeekToUndo(stockAndSalesByPlatform, weekNo);
        CellReference cellOfWeekToRemoveRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW - 1, columnToRemove - 1);
        weeklyReportSheet.getRow(cellOfWeekToRemoveRef.getRow()).getCell(cellOfWeekToRemoveRef.getCol()).setCellType(CellType.BLANK);
        for (int row = Constants.PLATFORM_HEADER_FIRST_ROW; row < Platforms.values().length + Constants.PLATFORM_HEADER_FIRST_ROW; row++) {
            CellReference cellToRemoveRef = new CellReference(row - 1, columnToRemove - 1);
            weeklyReportSheet.getRow(cellToRemoveRef.getRow()).getCell(cellToRemoveRef.getCol()).setCellType(CellType.BLANK);
        }
        //Checking if the lastly added week is the week to be removed.
        CellReference stockWeekNumberCellRef = new CellReference("BI" + Constants.PLATFORMS_TABLE_WEEK_ROW);
        Pattern pattern = Pattern.compile("w" + weekNo);
        Matcher m = pattern.matcher(weeklyReportSheet.getRow(stockWeekNumberCellRef.getRow()).getCell(stockWeekNumberCellRef.getCol()).getStringCellValue());
        if (m.matches()) {
            for (int row = Constants.PLATFORM_HEADER_FIRST_ROW; row < Platforms.values().length + Constants.PLATFORM_HEADER_FIRST_ROW; row++) {
                CellReference latestWeekStockCellRef = new CellReference("BI" + row);
                CellReference currentRowPlatformCellRef = new CellReference("B" + row);
                if (weeklyReportSheet.getRow(latestWeekStockCellRef.getRow()).getCell(latestWeekStockCellRef.getCol()).getCellTypeEnum() == CellType.BLANK) {
                    if (stockAndSalesByPlatform.get(weeklyReportSheet.getRow(currentRowPlatformCellRef
                            .getRow()).getCell(currentRowPlatformCellRef.getCol()).getStringCellValue()).Stock == 0) {
                        continue;
                    }
                } else {
                    Cell latestWeekStockCell = weeklyReportSheet.getRow(latestWeekStockCellRef.getRow()).getCell(latestWeekStockCellRef.getCol());
                    int stockParser;
                    if (NumberUtils.isParsable(latestWeekStockCell.getStringCellValue())) {
                        stockParser = Integer.parseInt(latestWeekStockCell.getStringCellValue());
                    } else {
                        break;
                    }
                    if (stockParser == stockAndSalesByPlatform.get(weeklyReportSheet.getRow(currentRowPlatformCellRef
                            .getRow()).getCell(currentRowPlatformCellRef.getCol()).getStringCellValue()).Stock) {
                        continue;
                    }
                    break;
                }
                for (int rowToDeleteOn = Constants.PLATFORM_HEADER_FIRST_ROW; rowToDeleteOn < Platforms.values().length + Constants.PLATFORM_HEADER_FIRST_ROW; rowToDeleteOn++) {
                    CellReference stockCellToBeRemoved = new CellReference("BI" + rowToDeleteOn);
                    weeklyReportSheet.getRow(stockCellToBeRemoved.getRow())
                            .getCell(stockCellToBeRemoved.getCol()).setCellValue(Constants.NO_DATA);
                }
            }
        }
    }

    private HashMap<String, StockSales> getStockSalesByPlatform() {
        HashMap<String, StockSales> stockSalesByPlatform = new HashMap<>();
        newData.entrySet().forEach((shop) -> {
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                if (!stockSalesByPlatform.containsKey(platform.getKey())) {
                    StockSales newSSObject = new StockSales();
                    newSSObject.Sales = Integer.MIN_VALUE;
                    newSSObject.Stock = Integer.MIN_VALUE;
                    stockSalesByPlatform.put(platform.getKey(), newSSObject);
                }
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    if (stockSalesByPlatform.get(platform.getKey()).Stock == Integer.MIN_VALUE
                            || stockSalesByPlatform.get(platform.getKey()).Sales == Integer.MIN_VALUE) {
                        stockSalesByPlatform.get(platform.getKey()).Stock = 0;
                        stockSalesByPlatform.get(platform.getKey()).Sales = 0;
                    }
                    stockSalesByPlatform.get(platform.getKey()).Stock += game.getValue().Stock;
                    stockSalesByPlatform.get(platform.getKey()).Sales += game.getValue().Sales;
                }
            }
        });
        return stockSalesByPlatform;
    }

    private int findWeekToUndo(HashMap<String, StockSales> newData, int weekNo) throws OutputFileNoRecordsFoundException {
        List<Integer> columnsMatchingWeeklyHeader = new ArrayList<>();
        boolean recordExists = false;
        for (int column = Constants.SELLOUT_TABLE_FIRST_COLUMN; column <= Constants.SELLOUT_TABLE_LAST_COLUMN; column++) {
            CellReference weeklyHeaderCellRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW - 1, column - 1);
            if (!"w".concat(Integer.toString(weekNo)).equals(weeklyReportSheet.getRow(weeklyHeaderCellRef.getRow())
                    .getCell(weeklyHeaderCellRef.getCol()).getStringCellValue())) {
                if (column == Constants.SELLOUT_TABLE_LAST_COLUMN && !recordExists) {
                    throw new OutputFileNoRecordsFoundException(rb.getString("OutputFileNoRecordsFoundExceptionMessage"));
                }
            } else if ("w".concat(Integer.toString(weekNo)).equals(weeklyReportSheet.getRow(weeklyHeaderCellRef.getRow())
                    .getCell(weeklyHeaderCellRef.getCol()).getStringCellValue())) {
                recordExists = true;
                columnsMatchingWeeklyHeader.add(column);
            }
        }
        for (int column = columnsMatchingWeeklyHeader.size() - 1; column >= 0; column--) {
            boolean continueSearching = false;
            for (int row = Constants.PLATFORM_HEADER_FIRST_ROW; row < Platforms.values().length + Constants.PLATFORM_HEADER_FIRST_ROW; row++) {
                CellReference currentPlatformCellRef = new CellReference("B" + row);
                CellReference currentValueCellRef = new CellReference(row - 1, columnsMatchingWeeklyHeader.get(column) - 1);
                String currentRowPlatform = weeklyReportSheet.getRow(currentPlatformCellRef.getRow()).getCell(currentPlatformCellRef.getCol()).getStringCellValue();
                if (weeklyReportSheet.getRow(currentValueCellRef.getRow()).getCell(currentValueCellRef.getCol()).getCellTypeEnum() == CellType.BLANK) {
                    if (newData.get(currentRowPlatform).Sales == Integer.MIN_VALUE) {
                        continue;
                    }
                    continueSearching = true;
                    break;
                } else if (weeklyReportSheet.getRow(currentValueCellRef.getRow()).getCell(currentValueCellRef.getCol()).getNumericCellValue()
                        != newData.get(weeklyReportSheet.getRow(currentPlatformCellRef.getRow()).getCell(currentPlatformCellRef.getCol()).getStringCellValue()).Sales) {
                    continueSearching = true;
                    break;
                }
            }
            if (column == 0 && continueSearching) {
                throw new OutputFileNoRecordsFoundException(rb.getString("OutputFileNoRecordsToBeUndoneFoundExceptionMessage"));
            }
            if (continueSearching == true) {
                continue;
            }
            return columnsMatchingWeeklyHeader.get(column);
        }
        return 0;
    }

    private void clearTopFiveStatistics() {
        for (int row = Constants.TOP_FIVE_TOP_FIRST_ROW; row <= Constants.TOP_FIVE_TOP_LAST_ROW; row++) {
            CellReference latestWeekShopCellRef = new CellReference("K" + row);
            CellReference latestWeekSalesCellRef = new CellReference("P" + row);
            topFiveSheet.getRow(latestWeekShopCellRef.getRow()).getCell(latestWeekShopCellRef.getCol()).setCellType(CellType.BLANK);
            topFiveSheet.getRow(latestWeekSalesCellRef.getRow()).getCell(latestWeekSalesCellRef.getCol()).setCellType(CellType.BLANK);
        }
        for (int row = Constants.TOP_FIVE_BOTTOM_FIRST_ROW; row <= Constants.TOP_FIVE_BOTTOM_LAST_ROW; row++) {
            CellReference latestWeekShopCellRef = new CellReference("K" + row);
            CellReference latestWeekSalesCellRef = new CellReference("P" + row);
            topFiveSheet.getRow(latestWeekShopCellRef.getRow()).getCell(latestWeekShopCellRef.getCol()).setCellType(CellType.BLANK);
            topFiveSheet.getRow(latestWeekSalesCellRef.getRow()).getCell(latestWeekSalesCellRef.getCol()).setCellType(CellType.BLANK);
        }
    }

    private boolean isOutputFileCorrect() {
        CellReference firstSheetLabelCellRef = new CellReference("A1");
        CellReference firstSheetStockLabelCellRef = new CellReference("BG1");
        if (outputWorkbook.getNumberOfSheets() != 4) {
            return false;
        } else if (!Constants.SELL_OUT.equals(weeklyReportSheet.getRow(firstSheetLabelCellRef.getRow()).getCell(firstSheetLabelCellRef.getCol()).getStringCellValue())) {
            return false;
        } else if (!Constants.STOCK.equals(weeklyReportSheet.getRow(firstSheetStockLabelCellRef.getRow()).getCell(firstSheetStockLabelCellRef.getCol()).getStringCellValue())) {
            return false;
        }
        return true;
    }

    private void writeTopFiveStatistics() {
        topFiveShopsBySalesOverall();
        topFiveGamesBySalesOverall();
        if (!undo) {
            topFiveShopsBySalesLatestWeek();
            topFiveGamesBySalesLatestWeek();
        }
    }

    private void topFiveShopsBySalesOverall() {
        HashMap<String, Integer> gamesAndSales = new HashMap<>();
        final int lastRowUsed = salesByPlatformSheet.getLastRowNum();
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row <= lastRowUsed; row++) {
            CellReference totalCellRef = new CellReference("P" + row);
            Cell totalCell = salesByPlatformSheet.getRow(totalCellRef.getRow()).getCell(totalCellRef.getCol());
            CellValue totalCellValue = evaluator.evaluate(totalCell);
            CellReference shopCellRef = new CellReference("A" + row);
            Cell shopCell = CellUtil.getRow(shopCellRef.getRow(), salesByPlatformSheet).getCell(shopCellRef.getCol());
            gamesAndSales.put(shopCell.getStringCellValue(), (int) totalCellValue.getNumberValue());
        }
        List<Entry<String, Integer>> sortedGamesAndSales = gamesAndSales.entrySet().stream().sorted(Entry.comparingByValue(Collections.reverseOrder())).collect(Collectors.toList());
        for (int row = Constants.TOP_FIVE_TOP_FIRST_ROW; row <= Constants.TOP_FIVE_TOP_LAST_ROW; row++) {
            if (gamesAndSales.size() > row - Constants.TOP_FIVE_TOP_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("C" + row);
                CellReference stockCellRef = new CellReference("H" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).setCellValue(sortedGamesAndSales.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellValue(sortedGamesAndSales.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getValue());
            }
        }
    }

    private void topFiveGamesBySalesOverall() {
        HashMap<String, HashMap<String, Integer>> platformsGamesAndSales = new HashMap<>();
        final int lastRowUsed = salesByGameSheet.getLastRowNum() + 1;
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row <= lastRowUsed; row++) {
            CellReference platformCellRef = new CellReference("A" + row);
            CellReference gameCellRef = new CellReference("B" + row);
            CellReference salesCellRef = new CellReference("E" + row);
            Cell platformCell = salesByGameSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
            Cell gameCell = salesByGameSheet.getRow(gameCellRef.getRow()).getCell(gameCellRef.getCol());
            Cell salesCell = salesByGameSheet.getRow(salesCellRef.getRow()).getCell(salesCellRef.getCol());
            if (platformsGamesAndSales.containsKey(platformCell.getStringCellValue())) {
                if (!platformsGamesAndSales.get(platformCell.getStringCellValue()).containsKey(gameCell.getStringCellValue())) {
                    platformsGamesAndSales.get(platformCell.getStringCellValue()).put(gameCell.getStringCellValue(), (int) salesCell.getNumericCellValue());
                }
            } else {
                platformsGamesAndSales.put(platformCell.getStringCellValue(), new HashMap<>());
                platformsGamesAndSales.get(platformCell.getStringCellValue()).put(gameCell.getStringCellValue(), (int) salesCell.getNumericCellValue());
            }
        }
        List<Entry<String, HashMap<String, Integer>>> gamesAndSalesList = platformsGamesAndSales.entrySet().stream().collect(Collectors.toList());
        HashMap<String, Integer> combinedPlatformsAndGames = new HashMap<>();
        for (Entry<String, HashMap<String, Integer>> platform : gamesAndSalesList) {
            for (Entry<String, Integer> game : platform.getValue().entrySet()) {
                combinedPlatformsAndGames.put(platform.getKey() + " " + game.getKey(), game.getValue());
            }
        }
        List<Entry<String, Integer>> sortedCombined = combinedPlatformsAndGames.entrySet().stream().sorted(Entry.comparingByValue(Collections.reverseOrder())).collect(Collectors.toList());

        for (int row = Constants.TOP_FIVE_BOTTOM_FIRST_ROW; row <= Constants.TOP_FIVE_BOTTOM_LAST_ROW; row++) {
            if (sortedCombined.size() > row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("C" + row);
                CellReference stockCellRef = new CellReference("H" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).setCellValue(sortedCombined.get(row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellValue(sortedCombined.get(row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW).getValue());
            }
        }
    }

    private void topFiveShopsBySalesLatestWeek() {
        HashMap<String, Integer> shopsAndSales = new HashMap<>();
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            if (!shopsAndSales.containsKey(shop.getKey())) {
                shopsAndSales.put(shop.getKey(), 0);
            }
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    shopsAndSales.put(shop.getKey(), shopsAndSales.get(shop.getKey()) + newData.get(shop.getKey()).get(platform.getKey()).get(game.getKey()).Sales);
                }
            }
        }
        List<Entry<String, Integer>> sortedShopsAndSalesList = shopsAndSales.entrySet().stream().sorted(Entry.comparingByValue(Collections.reverseOrder())).collect(Collectors.toList());
        for (int row = Constants.TOP_FIVE_TOP_FIRST_ROW; row <= Constants.TOP_FIVE_TOP_LAST_ROW; row++) {
            if (sortedShopsAndSalesList.size() > row - Constants.TOP_FIVE_TOP_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("K" + row);
                CellReference stockCellRef = new CellReference("P" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol())
                        .setCellValue(sortedShopsAndSalesList.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol())
                        .setCellValue(sortedShopsAndSalesList.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getValue());
            }
        }
    }

    private void topFiveGamesBySalesLatestWeek() {
        HashMap<String, Integer> platformsGamesAndSales = new HashMap<>();
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    if (platformsGamesAndSales.containsKey(platform.getKey() + " " + game.getKey())) {
                        platformsGamesAndSales.put(platform.getKey() + " " + game.getKey(),
                                platformsGamesAndSales.get(platform.getKey() + " " + game.getKey()) + game.getValue().Sales);
                    } else {
                        platformsGamesAndSales.put(platform.getKey() + " " + game.getKey(), game.getValue().Sales);
                    }
                }
            }
        }
        List<Entry<String, Integer>> sortedPlatformsGamesAndSales = platformsGamesAndSales.entrySet().stream().sorted(Entry.comparingByValue(Collections.reverseOrder())).collect(Collectors.toList());
        for (int row = Constants.TOP_FIVE_BOTTOM_FIRST_ROW; row <= Constants.TOP_FIVE_BOTTOM_LAST_ROW; row++) {
            if (sortedPlatformsGamesAndSales.size() > row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("K" + row);
                CellReference stockCellRef = new CellReference("P" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol())
                        .setCellValue(sortedPlatformsGamesAndSales.get(row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol())
                        .setCellValue(sortedPlatformsGamesAndSales.get(row - Constants.TOP_FIVE_BOTTOM_FIRST_ROW).getValue());
            }
        }
    }

    private void writeOverallSalesByPlatform() {
        if (salesByPlatformSheet.getLastRowNum() > 2) {
            overallSalesByPlatformExistingRecords();
        } else {
            overallSalesByPlatformFreshRecords();
        }
    }

    private void overallSalesByPlatformExistingRecords() {
        HashMap<String, HashMap<String, Integer>> currentStatistics = getCurrentOverallSalesPerPlatform();
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            if (!currentStatistics.containsKey(shop.getKey())) {
                currentStatistics.put(shop.getKey(), new HashMap<>());
            }
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                if (!currentStatistics.get(shop.getKey()).containsKey(platform.getKey())) {
                    currentStatistics.get(shop.getKey()).put(platform.getKey(), 0);
                }
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    if (!undo) {
                        currentStatistics.get(shop.getKey()).put(platform.getKey(), currentStatistics.get(shop.getKey()).get(platform.getKey()) + game.getValue().Sales);
                    } else {
                        currentStatistics.get(shop.getKey()).put(platform.getKey(), currentStatistics.get(shop.getKey()).get(platform.getKey()) - game.getValue().Sales);
                    }
                }
            }
        }
        List<Entry<String, HashMap<String, Integer>>> newStatistics = currentStatistics.entrySet().stream().collect(Collectors.toList());
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row < newStatistics.size() + Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row++) {
            CellRangeAddress shopNameCellAddress = CellRangeAddress.valueOf("A" + row + ":C" + row);
            try {
                salesByPlatformSheet.addMergedRegion(shopNameCellAddress);
            } catch (IllegalStateException ex) {
                //merged cells already exist. proceed
            }
            CellReference shopCellRef = new CellReference(row - 1, 0);
            Row shopRow = CellUtil.getRow(shopCellRef.getRow(), salesByPlatformSheet);
            Cell shopCell = salesByPlatformSheet.getRow(shopRow.getRowNum()).getCell(shopCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            shopCell.setCellValue(newStatistics.get(row - Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW).getKey());
            CellReference totalCellRef = new CellReference(row - 1, Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL - 1);
            Row totalRow = CellUtil.getRow(totalCellRef.getRow(), salesByPlatformSheet);
            Cell totalCell = salesByPlatformSheet.getRow(totalRow.getRowNum()).getCell(totalCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            totalCell.setCellType(CellType.FORMULA);
            totalCell.setCellFormula("SUM(D" + row + ":O" + row + ")");
            for (Entry<String, Integer> platform : newStatistics.get(row - Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW).getValue().entrySet()) {
                for (int column = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL; column < Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL; column++) {
                    CellReference platformCellRef = new CellReference(Constants.OVERALL_SALES_BY_PLATFORM_HEADER_ROW - 1, column - 1);
                    if (salesByPlatformSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol()).getStringCellValue().equals(platform.getKey())) {
                        CellReference platformSalesCellRef = new CellReference(row - 1, column - 1);
                        salesByPlatformSheet.getRow(platformSalesCellRef.getRow()).getCell(platformSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                                .setCellValue(platform.getValue());
                    }
                }
            }
        }
    }

    private void overallSalesByPlatformFreshRecords() {
        List<Entry<String, HashMap<String, HashMap<String, StockSales>>>> currentStatistics = newData.entrySet().stream().collect(Collectors.toList());
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row < currentStatistics.size() + Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row++) {
            CellRangeAddress shopNameCellAddress = CellRangeAddress.valueOf("A" + row + ":C" + row);
            salesByPlatformSheet.addMergedRegion(shopNameCellAddress);
            CellReference shopCellRef = new CellReference(row - 1, 0);
            Row shopRow = CellUtil.getRow(shopCellRef.getRow(), salesByPlatformSheet);
            Cell shopCell = salesByPlatformSheet.getRow(shopRow.getRowNum()).getCell(shopCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            shopCell.setCellValue(currentStatistics.get(row - Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW).getKey());
            CellReference totalCellRef = new CellReference(row - 1, Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL - 1);
            Row totalRow = CellUtil.getRow(totalCellRef.getRow(), salesByPlatformSheet);
            Cell totalCell = salesByPlatformSheet.getRow(totalRow.getRowNum()).getCell(totalCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            totalCell.setCellType(CellType.FORMULA);
            totalCell.setCellFormula("SUM(D" + row + ":O" + row + ")");

            for (Entry<String, HashMap<String, StockSales>> platform : currentStatistics.get(row - Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW).getValue().entrySet()) {
                int sumSales = 0;
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    sumSales += currentStatistics.get(row - Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW).getValue().get(platform.getKey()).get(game.getKey()).Sales;
                }
                for (int column = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL; column < Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL; column++) {
                    CellReference platformCellRef = new CellReference(Constants.OVERALL_SALES_BY_PLATFORM_HEADER_ROW - 1, column - 1);
                    if (salesByPlatformSheet.getRow(platformCellRef.getRow())
                            .getCell(platformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().equals(platform.getKey())) {
                        CellReference platformSalesCellRef = new CellReference(row - 1, column - 1);
                        salesByPlatformSheet.getRow(platformSalesCellRef.getRow())
                                .getCell(platformSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(sumSales);
                    }
                }
            }
        }
    }

    private HashMap<String, HashMap<String, Integer>> getCurrentOverallSalesPerPlatform() {
        HashMap<String, HashMap<String, Integer>> shopPlatformSales = new HashMap<>();
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row <= salesByPlatformSheet.getLastRowNum(); row++) {
            CellReference shopCellRef = new CellReference(row - 1, 0);
            Cell shopCell = salesByPlatformSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol());
            if ("".equals(shopCell.getStringCellValue())) {
                continue;
            }
            if (!shopPlatformSales.containsKey(shopCell.getStringCellValue())) {
                shopPlatformSales.put(shopCell.getStringCellValue(), new HashMap<>());
                for (int column = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL; column <= Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL; column++) {
                    CellReference platformCellRef = new CellReference(Constants.OVERALL_SALES_BY_PLATFORM_HEADER_ROW - 1, column - 1);
                    Cell platformCell = salesByPlatformSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
                    if (platformCell.getStringCellValue().equals(Constants.TOTAL)) {
                        break;
                    }
                    if (!shopPlatformSales.get(shopCell.getStringCellValue()).containsKey(platformCell.getStringCellValue())) {
                        CellReference platformSalesValueCellRef = new CellReference(row - 1, column - 1);
                        Cell platformSalesValueCell = CellUtil.getRow(platformSalesValueCellRef.getRow(), salesByPlatformSheet)
                                .getCell(platformSalesValueCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        if (platformSalesValueCell.getCellTypeEnum() != CellType.BLANK) {
                            shopPlatformSales.get(shopCell.getStringCellValue()).put(platformCell.getStringCellValue(), 0);
                            shopPlatformSales.get(shopCell.getStringCellValue())
                                    .put(platformCell.getStringCellValue(), shopPlatformSales.get(shopCell.getStringCellValue())
                                            .get(platformCell.getStringCellValue()) + (int) platformSalesValueCell.getNumericCellValue());
                        }
                    }
                }
            }
        }
        return shopPlatformSales;
    }

    private void writeOverallSalesByGame() {
        if (salesByGameSheet.getLastRowNum() > 2) {
            salesByGameExistingRecords();
        } else {
            salesByGameFreshRecords();
        }
    }

    private void salesByGameExistingRecords() {
        int currentLastRow;
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    //0 based + 1 to make it 1-based;
                    currentLastRow = salesByGameSheet.getLastRowNum() + 1;

                    for (int row = Constants.OVERALL_SALES_BY_GAME_FIRST_ROW; row <= currentLastRow; row++) {
                        CellReference platformCellRef = new CellReference("A" + row);
                        Cell platformCell = salesByGameSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
                        CellReference gameCellRef = new CellReference("B" + row);
                        Cell gameCell = salesByGameSheet.getRow(gameCellRef.getRow()).getCell(gameCellRef.getCol());
                        CellReference salesCellRef = new CellReference("E" + row);
                        Cell salesCell = salesByGameSheet.getRow(salesCellRef.getRow()).getCell(salesCellRef.getCol());
                        if (platformCell.getStringCellValue().equals(platform.getKey())) {
                            if (gameCell.getStringCellValue().equals(game.getKey())) {
                                if (!undo) {
                                    salesCell.setCellValue(salesCell.getNumericCellValue() + game.getValue().Sales);
                                } else {
                                    salesCell.setCellValue(salesCell.getNumericCellValue() - game.getValue().Sales);
                                }
                                break;
                            }
                        }
                        //If we've checked every row and neither was a match it doesnt exist
                        if (!undo && row == currentLastRow) {
                            CellReference nextRowPlatformCellRef = new CellReference("A" + (currentLastRow + 1));
                            CellReference nextRowGameCellRef = new CellReference("B" + (currentLastRow + 1));
                            CellReference nextRowSalesCellRef = new CellReference("E" + (currentLastRow + 1));
                            CellUtil.getRow(nextRowPlatformCellRef.getRow(), salesByGameSheet)
                                    .getCell(nextRowPlatformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(platform.getKey());
                            CellUtil.getRow(nextRowGameCellRef.getRow(), salesByGameSheet)
                                    .getCell(nextRowGameCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getKey());
                            CellUtil.getRow(nextRowSalesCellRef.getRow(), salesByGameSheet)
                                    .getCell(nextRowSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getValue().Sales);
                            CellRangeAddress gameCellsRange = CellRangeAddress.valueOf("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1));
                            salesByGameSheet.addMergedRegion(gameCellsRange);
                        }
                    }
                }
            }
        }
    }

    private void salesByGameFreshRecords() {
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    int currentLastRow = salesByGameSheet.getLastRowNum() + 1;
                    if (currentLastRow < Constants.OVERALL_SALES_BY_GAME_FIRST_ROW) {
                        CellReference nextRowPlatformCellRef = new CellReference("A" + (currentLastRow + 1));
                        CellReference nextRowGameCellRef = new CellReference("B" + (currentLastRow + 1));
                        CellReference nextRowSalesCellRef = new CellReference("E" + (currentLastRow + 1));
                        CellUtil.getRow(nextRowPlatformCellRef.getRow(), salesByGameSheet).getCell(nextRowPlatformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(platform.getKey());
                        CellUtil.getRow(nextRowGameCellRef.getRow(), salesByGameSheet).getCell(nextRowGameCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getKey());
                        CellUtil.getRow(nextRowSalesCellRef.getRow(), salesByGameSheet).getCell(nextRowSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getValue().Sales);
                        CellRangeAddress gameCellsRange = CellRangeAddress.valueOf("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1));
                        salesByGameSheet.addMergedRegion(gameCellsRange);
                        continue;
                    }
                    for (int row = Constants.OVERALL_SALES_BY_GAME_FIRST_ROW; row <= currentLastRow; row++) {
                        CellReference platformCellRef = new CellReference("A" + row);
                        Cell platformCell = salesByGameSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
                        CellReference gameCellRef = new CellReference("B" + row);
                        Cell gameCell = salesByGameSheet.getRow(gameCellRef.getRow()).getCell(gameCellRef.getCol());
                        CellReference salesCellRef = new CellReference("E" + row);
                        Cell salesCell = salesByGameSheet.getRow(salesCellRef.getRow()).getCell(salesCellRef.getCol());
                        if (platformCell.getStringCellValue().equals(platform.getKey())) {
                            if (gameCell.getStringCellValue().equals(game.getKey())) {
                                salesCell.setCellValue(salesCell.getNumericCellValue() + game.getValue().Sales);
                                break;
                            }
                        }
                        if (row == currentLastRow) {
                            CellReference nextRowPlatformCellRef = new CellReference("A" + (currentLastRow + 1));
                            CellReference nextRowGameCellRef = new CellReference("B" + (currentLastRow + 1));
                            CellReference nextRowSalesCellRef = new CellReference("E" + (currentLastRow + 1));
                            CellUtil.getRow(nextRowPlatformCellRef.getRow(), salesByGameSheet).getCell(nextRowPlatformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(platform.getKey());
                            CellUtil.getRow(nextRowGameCellRef.getRow(), salesByGameSheet).getCell(nextRowGameCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getKey());
                            CellUtil.getRow(nextRowSalesCellRef.getRow(), salesByGameSheet).getCell(nextRowSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(game.getValue().Sales);
                            CellRangeAddress gameCellsRange = CellRangeAddress.valueOf("B" + (currentLastRow + 1) + ":D" + (currentLastRow + 1));
                            salesByGameSheet.addMergedRegion(gameCellsRange);
                        }
                    }
                }
            }
        }
    }

    protected void formatDataHashMap() {
        for (Platforms platform : Platforms.values()) {
            newData.entrySet().stream().filter((store) -> (!store.getValue().containsKey(platform.getOutputAbbreviation()))).forEachOrdered((store) -> {
                store.getValue().put(platform.getOutputAbbreviation(), new HashMap<>());
            });
        }
    }
    //TODO abstract methods!

    /**
     * Get week number from the input file.
     *
     * @return week number.
     * @throws net.thecir.exceptions.InputFileContainsNoValidDateException
     */
    protected abstract int getWeekNumber() throws InputFileContainsNoValidDateException;

    protected abstract void readInputData();

    protected abstract boolean isInputFileCorrect();

    protected abstract String getStoreName(int row);

}
