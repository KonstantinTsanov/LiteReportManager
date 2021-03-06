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
package net.thecir.constants;

import net.thecir.enums.Platforms;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
public class Constants {

    public static final int PLATFORMS_TABLE_WEEK_ROW = 3;
    public static final int PLATFORMS_TABLE_LASTROW = Constants.PLATFORM_HEADER_FIRST_ROW + Platforms.values().length;
    public static final int PLATFORM_HEADER_FIRST_ROW = PLATFORMS_TABLE_WEEK_ROW + 1;
    public static final int PLATFORM_HEADER_LAST_ROW = PLATFORM_HEADER_FIRST_ROW - 1 + Platforms.values().length;

    public static final int SELLOUT_TABLE_FIRST_COLUMN = 3;
    public static final int SELLOUT_TABLE_LAST_COLUMN = 54;

    public static final int TOP_FIVE_TOP_FIRST_ROW = 7;
    public static final int TOP_FIVE_TOP_LAST_ROW = 11;
    public static final int TOP_FIVE_BOTTOM_FIRST_ROW = 16;
    public static final int TOP_FIVE_BOTTOM_LAST_ROW = 20;

    public static final int OVERALL_SALES_BY_PLATFORM_HEADER_ROW = 3;
    public static final int OVERALL_SALES_BY_PLATFORM_FIRST_ROW = 4;
    public static final int OVERALL_SALES_BY_PLATFORM_FIRST_COL = 4;
    //+1 for the total column, non 0-based
    public static final int OVERALL_SALES_BY_PLATFORM_LAST_COL = Platforms.values().length + OVERALL_SALES_BY_PLATFORM_FIRST_COL;

    public static final int OVERALL_SALES_BY_GAME_FIRST_ROW = 4;

    public static final String PLATFORMS_TABLE_TOTAL = "Total";
    public static final String TOTAL_PCS = "Total pcs";
    public static final String DAYS_IN_STOCK = "Days in stock";
    public static final String SALES = "Sales";
    public static final String STOCK = "Stock";
    public static final String GAME = "Game";
    public static final String SHOP = "Shop";
    public static final String TOTAL = "Total";
    public static final String PLATFORM = "Platform";
    public static final String SHEET_2_LABEL = "Top 5 statistics";
    public static final String TOP_LEFT_LABEL = "Top 5 shops by sales (Overall)";
    public static final String TOP_RIGHT_LABEL = "Top 5 shops by sales (Latest week)";
    public static final String BOTTOM_LEFT_LABEL = "Top 5 games by sales (Overall)";
    public static final String BOTTOM_RIGHT_LABEL = "Top 5 games by sales (Latest week)";

    public static final String SHEET_3_LABEL = "Overall sales by platform";
    public static final String SHEET_4_LABEL = "Overall sales by game";

    public static final String SHEET_2_NAME = "Top 5 statistics";
    public static final String SHEET_3_NAME = "Overall sales by platform";
    public static final String SHEET_4_NAME = "Overall sales by game";

    public static final String NAMCO = "Namco";
    public static final String SELL_OUT = "Sell out";

    //Must find a way to use an empty constant during new file creation, not hardcoding empty string.
    public static String NO_DATA = "";

}
