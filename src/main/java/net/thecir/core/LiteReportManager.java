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
package net.thecir.core;

import java.io.File;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Level;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import lombok.extern.java.Log;
import net.thecir.exceptions.NewFileCreationException;
import net.thecir.exceptions.OutputFileIOException;
import net.thecir.filemanagers.NewFileManager;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public class LiteReportManager {

    private final ExecutorService newFileExec = Executors.newFixedThreadPool(1);

    private static LiteReportManager SINGLETON;

    private JTextField outputField;

    public static LiteReportManager getInstance() {
        if (SINGLETON == null) {
            SINGLETON = new LiteReportManager();
        }
        return SINGLETON;
    }

    public void setOutputArea(JTextField field) {
        this.outputField = field;
    }

    public void createNewFile() {
        newFileExec.execute(() -> {
            try {
                File newFile = NewFileManager.getInstance().createNewWorkbook();
                if (newFile != null) {
                    outputField.setText(newFile.getAbsolutePath());
                }
            } catch (OutputFileIOException ex) {
                //TODO
                log.log(Level.SEVERE, null, ex);
                SwingUtilities.invokeLater(() -> {
                    throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
                });
            } catch (NewFileCreationException ex) {
                //TODO
                log.log(Level.SEVERE, null, ex);
                SwingUtilities.invokeLater(() -> {
                    throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
                });
            }
        });
    }
}
