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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map.Entry;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.java.Log;
import net.thecir.constants.Constants;
import net.thecir.enums.ExcelWorkbookType;
import net.thecir.enums.Platforms;
import net.thecir.exceptions.OutputFileIsFullException;
import net.thecir.exceptions.OutputFileNoRecordsFoundException;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;
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

    protected void writeToSheet() throws OutputFileIsFullException, OutputFileNoRecordsFoundException {
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
            CellReference weekCellRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW - 1, column - 1);
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

    private void undoWeeklyReport() throws OutputFileNoRecordsFoundException {
        int weekNo = getWeekNumber();
        HashMap<String, StockSales> stockAndSalesByPlatform = getStockSalesByPlatform();
        int columnToRemove = findWeekToUndo(stockAndSalesByPlatform, weekNo);
        CellReference cellOfWeekToRemoveRef = new CellReference(Constants.PLATFORMS_TABLE_WEEK_ROW - 1, columnToRemove - 1);
        weeklyReportSheet.getRow(cellOfWeekToRemoveRef.getRow()).getCell(cellOfWeekToRemoveRef.getCol()).setCellType(CellType.BLANK);
        for (int row = Constants.PLATFORMS_FIRSTROW; row < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; row++) {
            CellReference cellToRemoveRef = new CellReference(row - 1, columnToRemove - 1);
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
                for (int rowToDeleteOn = Constants.PLATFORMS_FIRSTROW; rowToDeleteOn < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; rowToDeleteOn++) {
                    CellReference stockCellToBeRemoved = new CellReference("BI" + rowToDeleteOn);
                    weeklyReportSheet.getRow(stockCellToBeRemoved.getRow()).getCell(stockCellToBeRemoved.getCol()).setCellValue(Constants.NO_DATA);
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
            if (!"w".concat(Integer.toString(weekNo)).equals(weeklyReportSheet.getRow(weeklyHeaderCellRef.getRow()).getCell(weeklyHeaderCellRef.getCol()).getStringCellValue())) {
                if (column == Constants.SELLOUT_TABLE_LAST_COLUMN && !recordExists) {
                    throw new OutputFileNoRecordsFoundException(rb.getString("OutputFileNoRecordsFoundExceptionMessage"));
                }
            } else if ("w".concat(Integer.toString(weekNo)).equals(weeklyReportSheet.getRow(weeklyHeaderCellRef.getRow()).getCell(weeklyHeaderCellRef.getCol()).getStringCellValue())) {
                recordExists = true;
                columnsMatchingWeeklyHeader.add(column);
            }
        }
        for (int column = columnsMatchingWeeklyHeader.size() - 1; column >= 0; column--) {
            boolean continueSearching = false;
            for (int row = Constants.PLATFORMS_FIRSTROW; row < Platforms.values().length + Constants.PLATFORMS_FIRSTROW; row++) {
                CellReference currentPlatformCellRef = new CellReference("B" + row);
                CellReference currentValueCellRef = new CellReference(row - 1, columnsMatchingWeeklyHeader.get(column) - 1);
                String currentRowPlatform = weeklyReportSheet.getRow(currentPlatformCellRef.getRow()).getCell(currentPlatformCellRef.getCol()).getStringCellValue();
                if (weeklyReportSheet.getRow(currentValueCellRef.getRow()).getCell(currentValueCellRef.getCol()).getCellTypeEnum() == CellType.BLANK) {
                    if (newData.get(currentRowPlatform).Sales == Integer.MIN_VALUE) {
                        continue;
                    }
                    continueSearching = true;
                    break;
                } else if (Integer.parseInt(weeklyReportSheet.getRow(currentValueCellRef.getRow()).getCell(currentValueCellRef.getCol()).getStringCellValue())
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

    private boolean outputFileSignature() {
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
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row <= salesByPlatformSheet.getLastRowNum() + 1; row++) {
            CellReference totalCellRef = new CellReference("P" + row);
            int num;
            Cell totalCell = salesByPlatformSheet.getRow(totalCellRef.getRow()).getCell(totalCellRef.getCol());
            if (NumberUtils.isParsable(totalCell.getStringCellValue())) {
                num = Integer.parseInt(totalCell.getStringCellValue());
            } else {
                continue;
            }
            CellReference shopCellRef = new CellReference("A" + row);
            gamesAndSales.put(salesByPlatformSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).getStringCellValue(), num);
        }
        List<Entry<String, Integer>> sortedGamesAndSales = gamesAndSales.entrySet().stream().sorted(Entry.comparingByValue()).collect(Collectors.toList());
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
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row <= salesByGameSheet.getLastRowNum(); row++) {
            CellReference platformCellRef = new CellReference("A" + row);
            CellReference gameCellRef = new CellReference("B" + row);
            CellReference salesCellRef = new CellReference("E" + row);
            Cell platformCell = salesByGameSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol());
            Cell gameCell = salesByGameSheet.getRow(gameCellRef.getRow()).getCell(gameCellRef.getCol());
            Cell salesCell = salesByGameSheet.getRow(salesCellRef.getRow()).getCell(salesCellRef.getCol());
            if (platformsGamesAndSales.containsKey(platformCell.getStringCellValue())) {
                if (!platformsGamesAndSales.get(platformCell.getStringCellValue()).containsKey(gameCell.getStringCellValue())) {
                    int sales;
                    if (NumberUtils.isParsable(salesCell.getStringCellValue())) {
                        sales = Integer.parseInt(salesCell.getStringCellValue());
                        platformsGamesAndSales.get(platformCell.getStringCellValue()).put(gameCell.getStringCellValue(), sales);
                    }
                }
            } else {
                platformsGamesAndSales.put(platformCell.getStringCellValue(), new HashMap<>());
                platformsGamesAndSales.get(platformCell.getStringCellValue()).put(gameCell.getStringCellValue(), Integer.parseInt(salesCell.getStringCellValue()));
            }
        }
        List<Entry<String, HashMap<String, Integer>>> gamesAndSalesList = platformsGamesAndSales.entrySet().stream().collect(Collectors.toList());
        HashMap<String, Integer> combinedPlatformsAndGames = new HashMap<>();
        for (Entry<String, HashMap<String, Integer>> platform : gamesAndSalesList) {
            for (Entry<String, Integer> game : platform.getValue().entrySet()) {
                combinedPlatformsAndGames.put(platform.getKey() + " " + game.getKey(), game.getValue());
            }
        }
        List<Entry<String, Integer>> sortedCombined = combinedPlatformsAndGames.entrySet().stream().sorted(Entry.comparingByValue()).collect(Collectors.toList());

        for (int row = Constants.TOP_FIVE_BOTTOM_FIRST_ROW; row <= Constants.TOP_FIVE_BOTTOM_LAST_ROW; row++) {
            if (sortedCombined.size() > row - Constants.TOP_FIVE_TOP_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("C" + row);
                CellReference stockCellRef = new CellReference("H" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).setCellValue(sortedCombined.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellValue(sortedCombined.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getValue());
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
        List<Entry<String, Integer>> sortedShopsAndSalesList = shopsAndSales.entrySet().stream().sorted(Entry.comparingByValue()).collect(Collectors.toList());
        for (int row = Constants.TOP_FIVE_TOP_FIRST_ROW; row <= Constants.TOP_FIVE_TOP_LAST_ROW; row++) {
            if (sortedShopsAndSalesList.size() > row - Constants.TOP_FIVE_TOP_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("K" + row);
                CellReference stockCellRef = new CellReference("P" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).setCellValue(sortedShopsAndSalesList.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellValue(sortedShopsAndSalesList.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getValue());
            }
        }
    }

    private void topFiveGamesBySalesLatestWeek() {
        HashMap<String, Integer> platformsGamesAndSales = new HashMap<>();
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            for (Entry<String, HashMap<String, StockSales>> platform : shop.getValue().entrySet()) {
                for (Entry<String, StockSales> game : platform.getValue().entrySet()) {
                    if (platformsGamesAndSales.containsKey(platform.getKey() + " " + game.getKey())) {
                        platformsGamesAndSales.put(platform.getKey() + " " + game.getKey(), platformsGamesAndSales.get(platform.getKey() + " " + game.getKey()) + game.getValue().Sales);
                    } else {
                        platformsGamesAndSales.put(platform.getKey() + " " + game.getKey(), game.getValue().Sales);
                    }
                }
            }
        }
        List<Entry<String, Integer>> sortedPlatformsGamesAndSales = platformsGamesAndSales.entrySet().stream().sorted(Entry.comparingByValue()).collect(Collectors.toList());
        for (int row = Constants.TOP_FIVE_BOTTOM_FIRST_ROW; row <= Constants.TOP_FIVE_BOTTOM_LAST_ROW; row++) {
            if (sortedPlatformsGamesAndSales.size() > row - Constants.TOP_FIVE_TOP_FIRST_ROW) {
                CellReference shopCellRef = new CellReference("K" + row);
                CellReference stockCellRef = new CellReference("P" + row);
                topFiveSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol()).setCellValue(sortedPlatformsGamesAndSales.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getKey());
                topFiveSheet.getRow(stockCellRef.getRow()).getCell(stockCellRef.getCol()).setCellValue(sortedPlatformsGamesAndSales.get(row - Constants.TOP_FIVE_TOP_FIRST_ROW).getValue());
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
        HashMap<String, HashMap<String, Integer>> currentStatistics = getCurrentOverallSalesByPlatform();
        for (Entry<String, HashMap<String, HashMap<String, StockSales>>> shop : newData.entrySet()) {
            if (!currentStatistics.containsKey(shop.getKey())) {
                currentStatistics.put(shop.getKey(), new HashMap<>());
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
        }
        List<Entry<String, HashMap<String, Integer>>> newStatistics = currentStatistics.entrySet().stream().collect(Collectors.toList());
        for (int row = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row < newStatistics.size() + Constants.OVERALL_SALES_BY_PLATFORM_FIRST_ROW; row++) {
            CellRangeAddress shopNameCellAddress = CellRangeAddress.valueOf("A" + row + ":C" + row);
            salesByPlatformSheet.addMergedRegion(shopNameCellAddress);
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
                for (int column = Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL; column < Constants.OVERALL_SALES_BY_PLATFORM_LAST_COL; row++) {
                    CellReference platformCellRef = new CellReference(Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL - 1, column - 1);
                    if (salesByPlatformSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().equals(platform.getKey())) {
                        CellReference platformSalesCellRef = new CellReference(row - 1, column - 1);
                        salesByPlatformSheet.getRow(platformSalesCellRef.getRow()).getCell(platformSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(platform.getValue());
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
                    CellReference platformCellRef = new CellReference(Constants.OVERALL_SALES_BY_PLATFORM_FIRST_COL - 1, column - 1);
                    if (salesByPlatformSheet.getRow(platformCellRef.getRow()).getCell(platformCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).getStringCellValue().equals(platform.getKey())) {
                        CellReference platformSalesCellRef = new CellReference(row - 1, column - 1);
                        salesByPlatformSheet.getRow(platformSalesCellRef.getRow()).getCell(platformSalesCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(sumSales);
                    }
                }
            }
        }
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
