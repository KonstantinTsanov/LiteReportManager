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

import com.thecir.tools.ExcelTools;
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
import net.thecir.constants.TechnopolisConstants;
import net.thecir.enums.Platforms;
import net.thecir.exceptions.InputFileContainsNoValidDateException;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellUtil;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public class TechnopolisReportManager extends ReportManager {

    public TechnopolisReportManager(File inputFilePath, File outputFilePath, boolean undo) {
        super(inputFilePath, outputFilePath, undo);
    }

    @Override
    protected void formatDataHashMap() {
        String store;
        for (int row = TechnopolisConstants.FIRST_ROW; row <= inputDataSheet.getLastRowNum(); row++) {
            store = getStoreName(row);
            if (store != null) {
                if (!newData.containsKey(store)) {
                    newData.put(store, new HashMap<>());
                }
            }
        }
        super.formatDataHashMap();
    }

    @Override
    protected int getWeekNumber() throws InputFileContainsNoValidDateException {
        try {
            Date date = getDate();
            if (date == null) {
                throw new InputFileContainsNoValidDateException("The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
            }
            Calendar cal = Calendar.getInstance();
            cal.setTime(date);
            return cal.get(Calendar.WEEK_OF_YEAR);
        } catch (ParseException ex) {
            log.log(Level.SEVERE, "Unparsable source file date!", ex);
            throw new InputFileContainsNoValidDateException("The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
        }
    }

    /**
     * Helper method, because the file seems too chaotic and the date looks
     * placed there by hand. Therefore it could be placed in any of the 3 cells
     * at the top of the sheet.
     *
     * @return The exact cell, which contains the date.
     * @throws InputFileContainsNoValidDateException thrown if no parsable date
     * was found in any of the cells A1, B1 or C1.
     * @throws java.text.ParseException
     */
    protected Date getDate() throws InputFileContainsNoValidDateException, ParseException {
        for (int column = 0; column < 3; column++) {
            CellReference possibleDateCellRef = new CellReference(0, column);
            Cell possibleDateCell = inputDataSheet.getRow(possibleDateCellRef.getRow()).getCell(possibleDateCellRef.getCol());
            if (possibleDateCell.getCellTypeEnum() != CellType.STRING
                    || "".equals(possibleDateCell.getStringCellValue())) {
                if (column == 2) {
                    throw new InputFileContainsNoValidDateException(
                            "Please add \"From-To\" dates in any of the following cells: A1, B1, C1. The date format should be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY.");
                }
                continue;
            }
            String dateStringWithoutSpaces = possibleDateCell.getStringCellValue().replaceAll("\\s", "");
            Pattern pattern = Pattern.compile("-\\d{2}\\.\\d{2}\\.\\d{2}(?:\\d{2})?");
            boolean matchedOnce = false;
            Matcher matcher = pattern.matcher(dateStringWithoutSpaces);
            String extractedDate = null;
            while (matcher.find()) {
                if (matchedOnce == true) {
                    throw new InputFileContainsNoValidDateException(
                            "The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
                }
                matchedOnce = true;
                extractedDate = matcher.group().substring(1);
            }
            if (extractedDate == null) {
                continue;
            }
            SimpleDateFormat parser;
            if (extractedDate.length() == 10) {
                parser = new SimpleDateFormat("dd.MM.yyyy");
            } else {
                parser = new SimpleDateFormat("dd.MM.yy");
            }
            return parser.parse(extractedDate);
        }
        return null;
    }

    @Override
    protected void readInputData() {
        final int lastRowUsed = inputDataSheet.getLastRowNum();
        for (int row = TechnopolisConstants.FIRST_ROW; row <= lastRowUsed; row++) {
            CellReference itemNumberCellRef = new CellReference(row, TechnopolisConstants.ITEM_COLUMN);
            CellReference gameDescriptionCellRef = new CellReference(row, TechnopolisConstants.GAME_DESCR_COLUMN);
            CellReference soldQuantityCellRef = new CellReference(row, TechnopolisConstants.SOLD_QUANTITY_COLUMN);
            CellReference stockCellRef = new CellReference(row, TechnopolisConstants.STOCK_COLUMN);
            CellReference nextRowItemNumberCellRef;

            Cell itemNumberCell = CellUtil.getRow(itemNumberCellRef.getRow(), inputDataSheet).getCell(itemNumberCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell gameDescriptionCell = CellUtil.getRow(gameDescriptionCellRef.getRow(), inputDataSheet).getCell(gameDescriptionCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell soldQuantityCell = CellUtil.getRow(soldQuantityCellRef.getRow(), inputDataSheet).getCell(soldQuantityCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell stockCell = CellUtil.getRow(stockCellRef.getRow(), inputDataSheet).getCell(stockCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            Cell nextRowItemNumberCell;

            if ("".equals(ExcelTools.getStringCellValue(itemNumberCell)) || !NumberUtils.isParsable(ExcelTools.getStringCellValue(itemNumberCell))) {
                continue;
            }
            String currentPlatform = null;
            String currentTitle = null;
            for (Platforms platform : Platforms.values()) {
                Pattern pattern = Pattern.compile("^" + platform.getTechnopolisAbbreviation(), Pattern.CASE_INSENSITIVE);
                Matcher matcher = pattern.matcher(ExcelTools.getStringCellValue(gameDescriptionCell).trim());
                if ("".equals(ExcelTools.getStringCellValue(gameDescriptionCell)) || !matcher.find()) {
                    if (platform.ordinal() == (Platforms.Other.ordinal() - 1)) {
                        currentPlatform = Platforms.Other.name();
                        //Trimming the last character;
                        currentTitle = gameDescriptionCell.getStringCellValue().trim().substring(0, gameDescriptionCell.getStringCellValue().length() - 1);
                        break;
                    }
                    continue;
                }
                currentPlatform = platform.getOutputAbbreviation();
                currentTitle = gameDescriptionCell.getStringCellValue().trim()
                        .substring(0, gameDescriptionCell.getStringCellValue().length() - 1).replaceAll("^" + platform.getTechnopolisAbbreviation(), "").trim();
                break;
            }
            do {
                String store = getStoreName(row);
                if (!newData.get(store).get(currentPlatform).containsKey(currentTitle)) {
                    newData.get(store).get(currentPlatform).put(currentTitle, new StockSales());
                }
                if (NumberUtils.isParsable(ExcelTools.getStringCellValue(stockCell))) {
                    double stock = Double.parseDouble(ExcelTools.getStringCellValue(stockCell));
                    StockSales updatedStockSales = newData.get(store).get(currentPlatform).get(currentTitle);
                    updatedStockSales.Stock += (int) stock;
                    newData.get(store).get(currentPlatform).put(currentTitle, updatedStockSales);
                }
                if (NumberUtils.isParsable(ExcelTools.getStringCellValue(soldQuantityCell))) {
                    double sales = Double.parseDouble(ExcelTools.getStringCellValue(soldQuantityCell));
                    StockSales updatedStockSales = newData.get(store).get(currentPlatform).get(currentTitle);
                    updatedStockSales.Sales += (int) sales;
                    newData.get(store).get(currentPlatform).put(currentTitle, updatedStockSales);
                }
                row++;

                itemNumberCellRef = new CellReference(row, TechnopolisConstants.ITEM_COLUMN);
                nextRowItemNumberCellRef = new CellReference(row + 1, TechnopolisConstants.ITEM_COLUMN);
                soldQuantityCellRef = new CellReference(row, TechnopolisConstants.SOLD_QUANTITY_COLUMN);
                stockCellRef = new CellReference(row, TechnopolisConstants.STOCK_COLUMN);

                itemNumberCell = CellUtil.getRow(itemNumberCellRef.getRow(), inputDataSheet).getCell(itemNumberCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                nextRowItemNumberCell = CellUtil.getRow(nextRowItemNumberCellRef.getRow(), inputDataSheet).getCell(nextRowItemNumberCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                soldQuantityCell = CellUtil.getRow(soldQuantityCellRef.getRow(), inputDataSheet).getCell(soldQuantityCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                stockCell = CellUtil.getRow(stockCellRef.getRow(), inputDataSheet).getCell(stockCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

            } while (row != lastRowUsed && "".equals(ExcelTools.getStringCellValue(itemNumberCell)) && "".equals(ExcelTools.getStringCellValue(nextRowItemNumberCell)));
        }
    }

    @Override
    protected boolean isInputFileCorrect() {
        Pattern pt = Pattern.compile("^(Технополис|Видеолукс|WEB|GSM)");
        for (int row = TechnopolisConstants.FIRST_ROW; row <= inputDataSheet.getLastRowNum(); row++) {
            CellReference shopCellRef = new CellReference("C" + (row + 1));
            Cell shopCell = CellUtil.getRow(shopCellRef.getRow(), inputDataSheet).getCell(shopCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (shopCell.getCellTypeEnum() != CellType.STRING) {
                continue;
            }
            Matcher m = pt.matcher(shopCell.getStringCellValue());
            if (m.find()) {
                return true;
            }
        }
        return false;
    }

    @Override
    protected String getStoreName(int row) {
        //Matches only if the string does not begin with any of the strings in the braces and has one or more symbols (.+). Therefore if its an empty string it wont match.
        Pattern pt = Pattern.compile("^(?!Обект|Резултат|\\s).+");
        CellReference shopCellRef = new CellReference(row, TechnopolisConstants.SHOP_COLUMN);
        Row shopRow = CellUtil.getRow(shopCellRef.getRow(), inputDataSheet);
        Cell shopCell = shopRow.getCell(shopCellRef.getCol(), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        Matcher m = pt.matcher(shopCell.getStringCellValue().trim());
        if (m.find()) {
            return shopCell.getStringCellValue().trim();
        }
        return null;
    }
}
