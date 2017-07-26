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
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.extern.java.Log;
import net.thecir.constants.TechnopolisConstants;
import net.thecir.exceptions.InputFileContainsNoValidDateException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;

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
            store = getShopOnRow(row);
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
        Cell dateCell = getCellContainingDate();
        if (dateCell != null) {
            String dateStringWithoutSpaces = dateCell.getStringCellValue().replaceAll("\\s", "");
            Pattern pattern = Pattern.compile("-\\d{2}\\.\\d{2}\\.\\d{2}(?:\\d{2})?");
            boolean matchedOnce = false;
            Matcher matcher = pattern.matcher(dateStringWithoutSpaces);
            String extractedDate = null;
            while (matcher.find()) {
                if (matchedOnce == true) {
                    throw new InputFileContainsNoValidDateException("The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
                }
                matchedOnce = true;
                extractedDate = matcher.group().substring(1);
            }
            SimpleDateFormat parser;
            if (extractedDate.length() == 10) {
                parser = new SimpleDateFormat("dd.mm.yyyy");
            } else {
                parser = new SimpleDateFormat("dd.mm.yy");
            }
            try {
                Date date = parser.parse(extractedDate);
                Calendar cal = Calendar.getInstance();
                cal.setTime(date);
                return cal.get(Calendar.WEEK_OF_YEAR);
            } catch (ParseException ex) {
                log.log(Level.SEVERE, "Unparsable date!", ex);
                throw new InputFileContainsNoValidDateException("The date format must be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY. The date must be located in any of the following cells: A1, B1, C1.");
            }
        }
        return -1;
    }

    private Cell getCellContainingDate() throws InputFileContainsNoValidDateException {
        for (int column = 0; column < 3; column++) {
            CellReference possibleDateCellRef = new CellReference(0, column);
            Cell possibleDateCell = inputDataSheet.getRow(possibleDateCellRef.getRow()).getCell(possibleDateCellRef.getCol());
            if (possibleDateCell.getCellTypeEnum() == CellType.BLANK) {
                if (column == 2) {
                    throw new InputFileContainsNoValidDateException("Please add \"From-To\" dates in any of the following cells: A1, B1, C1. The date format should be DD.MM-DD.MM.YY or DD.MM-DD.MM.YYYY.");
                }
                continue;
            }
            return possibleDateCell;
        }
        return null;
    }

    @Override
    protected void readInputData() {
        for (int row = TechnopolisConstants.FIRST_ROW; row <= inputDataSheet.getLastRowNum(); row++) {
            
        }
    }

    @Override
    protected boolean isInputFileCorrect() {
        Pattern pt = Pattern.compile("^(Технополис|Видеолукс|WEB|GSM)");
        for (int row = TechnopolisConstants.FIRST_ROW; row <= inputDataSheet.getLastRowNum(); row++) {
            CellReference shopCellRef = new CellReference("C" + (row + 1));
            Cell shopCell = inputDataSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol());
            Matcher m = pt.matcher(shopCell.getStringCellValue());
            if (m.find()) {
                return true;
            }
        }
        return false;
    }

    @Override
    protected String getShopOnRow(int row) {
        Pattern pt = Pattern.compile("^(?!Обект|Резултат|\\s).*");
        CellReference shopCellRef = new CellReference(row, TechnopolisConstants.SHOP_COLUMN);
        Cell shopCell = inputDataSheet.getRow(shopCellRef.getRow()).getCell(shopCellRef.getCol());
        Matcher m = pt.matcher(shopCell.getStringCellValue());
        if (m.find()) {
            return shopCell.getStringCellValue();
        }
        return null;
    }
}
