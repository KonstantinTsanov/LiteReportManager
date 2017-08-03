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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.extern.java.Log;
import net.thecir.constants.TechnomarketConstants;
import net.thecir.enums.Platforms;
import net.thecir.exceptions.InputFileContainsNoValidDateException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public class TechnomarketReportManager extends ReportManager {

    public TechnomarketReportManager(File inputFilePath, File outputFilePath, boolean undo) {
        super(inputFilePath, outputFilePath, undo);
    }

    @Override
    protected void formatDataHashMap() {
        final int lastColumnUsed = inputDataSheet.getRow(TechnomarketConstants.SHOPS_ROW).getLastCellNum() - 1;//1-based number, apache pls.... -1 to make it 0-based again
        for (int column = TechnomarketConstants.SHOPS_FIRST_COLUMN; column <= lastColumnUsed; column++) {
            String store = getStoreName(column);
            if (!"".equals(store)) {
                if (!newData.containsKey(store)) {
                    newData.put(store, new HashMap<>());
                }
            }
        }
        super.formatDataHashMap();
    }

    @Override
    protected int getWeekNumber() throws InputFileContainsNoValidDateException {
        Date[] dates = new Date[2];
        Pattern pt = Pattern.compile("\\d{2}\\.\\d{2}\\.\\d{4}");
        Cell infoCell = inputDataSheet.getRow(TechnomarketConstants.INFO_CELL_ROW).getCell(TechnomarketConstants.INFO_CELL_COL);
        Matcher m = pt.matcher(infoCell.getStringCellValue());
        int datesCount = 0;
        while (m.find()) {
            if (datesCount >= 2) {
                throw new InputFileContainsNoValidDateException(
                        rb.getString("TechnopolisInputNoValidDate"));
            }
            String extractedDate = m.group(0);
            SimpleDateFormat parser = new SimpleDateFormat("dd.MM.yyyy");
            try {
                dates[datesCount] = parser.parse(extractedDate);
                datesCount++;
            } catch (ParseException ex) {
                log.log(Level.SEVERE, "Unparsable source file date!", ex);
                throw new InputFileContainsNoValidDateException("TechnopolisInputNoValidDate");
            }
        }
        if (dates[0].compareTo(dates[1]) > 0) {
            Date buffer = dates[0];
            dates[0] = dates[1];
            dates[1] = buffer;
        }
        Calendar cal = Calendar.getInstance();
        //The greater date
        cal.setTime(dates[1]);
        return cal.get(Calendar.WEEK_OF_YEAR);
    }

    @Override
    protected void readInputData() {
        int lastRow = inputDataSheet.getLastRowNum();
        for (int row = TechnomarketConstants.SHEET_FIRST_ROW; row <= lastRow; row++) {
            CellReference gameNumberCellRef = new CellReference("C" + (row + 1));
            Cell gameNumberCell = CellUtil.getRow(gameNumberCellRef.getRow(), inputDataSheet).getCell(gameNumberCellRef.getCol());
            if (gameNumberCell == null || gameNumberCell.getCellTypeEnum() != CellType.NUMERIC || (String.valueOf((long) gameNumberCell.getNumericCellValue()).length() != 12
                    && String.valueOf((long) gameNumberCell.getNumericCellValue()).length() != 13)) {
                continue;
            }
            String gamePlatform = "";
            String gameTitle = "";
            CellReference platformAndGameCellRef = new CellReference("B" + (row + 1));
            Cell platformAndGameCell = inputDataSheet.getRow(platformAndGameCellRef.getRow()).getCell(platformAndGameCellRef.getCol());
            for (Platforms platform : Platforms.values()) {
                Pattern pt = Pattern.compile("^" + platform.getTechnomarketAbbreviation(), Pattern.CASE_INSENSITIVE);
                String platformAndGame = platformAndGameCell.getStringCellValue().trim();
                //The fist space is removed, its the space between xbox and 360 -> xbox 360 becomes xbox360
                int indexOfFirstSpace = platformAndGame.indexOf(" ");
                if (indexOfFirstSpace >= 0) {
                    platformAndGame = platformAndGame.substring(0, indexOfFirstSpace) + "" + platformAndGame.substring(indexOfFirstSpace + 1);
                }
                Matcher m = pt.matcher(platformAndGame);
                if (m.find()) {
                    gamePlatform = platform.getOutputAbbreviation();
                    gameTitle = platformAndGame.replaceAll("^" + platform.getTechnomarketAbbreviation(), "").trim();
                    break;
                }
            }
            if ("".equals(gamePlatform)) {
                gamePlatform = Platforms.Other.getOutputAbbreviation();
                gameTitle = platformAndGameCell.getStringCellValue();
            }

            int lastColumnUsed = inputDataSheet.getRow(TechnomarketConstants.SHOPS_ROW).getLastCellNum() - 1;//1-based number, apache pls.... -1 to make it 0-based again
            for (int column = TechnomarketConstants.SHOPS_FIRST_COLUMN; column <= lastColumnUsed; column++) {
                String store = getStoreName(column);
                if ("".equals(store)) {
                    //exception??
                }
                if (!newData.get(store).get(gamePlatform).containsKey(gameTitle)) {
                    //adding each game, along with the stock and sales
                    //1d stock, 2d sales
                    newData.get(store).get(gamePlatform).put(gameTitle, new StockSales());
                }
                CellReference dataCellReference = new CellReference(row, column);
                Cell dataCell = inputDataSheet.getRow(dataCellReference.getRow()).getCell(dataCellReference.getCol());
                //stock
                if (column % 2 == 1) {
                    StockSales currentStockSales = newData.get(store).get(gamePlatform).get(gameTitle);
                    currentStockSales.Stock += (int) dataCell.getNumericCellValue();
                    newData.get(store).get(gamePlatform).put(gameTitle, currentStockSales);
                }/*sales*/ else {
                    StockSales currentStockSales = newData.get(store).get(gamePlatform).get(gameTitle);
                    currentStockSales.Sales += (int) dataCell.getNumericCellValue();
                    newData.get(store).get(gamePlatform).put(gameTitle, currentStockSales);
                }
            }
        }
    }

    @Override
    protected boolean isInputFileCorrect() {
        Pattern pt = Pattern.compile("technomarket", Pattern.CASE_INSENSITIVE);
        Cell infoCell = inputDataSheet.getRow(TechnomarketConstants.INFO_CELL_ROW).getCell(TechnomarketConstants.INFO_CELL_COL);
        if (infoCell == null || infoCell.getCellTypeEnum() != CellType.STRING) {
            return false;
        }
        Matcher m = pt.matcher(infoCell.getStringCellValue());
        boolean found = m.find();
        return found;
    }

    @Override
    protected String getStoreName(int column) {
        CellReference shopCellRef = new CellReference(TechnomarketConstants.SHOPS_ROW, column);
        Cell shopCell = inputDataSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol());
        String shop = shopCell.getStringCellValue();
        shop = shop.replaceAll("^(?:\\d*)?", "").trim();
        return shop;
    }
}
