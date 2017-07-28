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
package net.thecir.enums;

import lombok.Getter;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
public enum Platforms {
    PS2("PS2", "P2", "PS2"),
    PS3("PS3", "P3", "PS3"),
    PS4("PS4", "P4", "PS4"),
    XBOX360("XBOX360", "XB3", "XBOX360"),
    XBOXONE("XBOXONE", "XBO", "XBOXONE"),
    WII("WII", "WII", "WII"),
    PSP("PSP", "PSP", "PSP"),
    DS3("3DS", "3D", "3DS"),
    PSVITA("PSVITA", "PSV", "PSVITA"),
    PC("PC", "PC", "PC"),
    NDS("NDS", "DS", "NDS"),
    Other("Other", "Other", "Other");
    @Getter
    private final String outputAbbreviation;
    @Getter
    private final String technopolisAbbreviation;
    @Getter
    private final String technomarketAbbreviation;

    private Platforms(String outputAbbreviation, String technopolisAbbreviation, String technomarketAbbreviation) {
        this.outputAbbreviation = outputAbbreviation;
        this.technopolisAbbreviation = technopolisAbbreviation;
        this.technomarketAbbreviation = technomarketAbbreviation;
    }
}
