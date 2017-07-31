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
import java.util.Locale;
import java.util.ResourceBundle;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Level;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import lombok.extern.java.Log;
import net.thecir.enums.Stores;
import net.thecir.exceptions.InputFileContainsNoValidDateException;
import net.thecir.exceptions.InputFileNotMatchingSelectedFileException;
import net.thecir.exceptions.NewFileCreationException;
import net.thecir.exceptions.OutputFileIOException;
import net.thecir.exceptions.OutputFileIsFullException;
import net.thecir.exceptions.OutputFileNoRecordsFoundException;
import net.thecir.exceptions.OutputFileNotCorrectException;
import net.thecir.filemanagers.NewFileManager;
import net.thecir.reportmanagers.ReportManager;
import net.thecir.reportmanagers.TechnomarketReportManager;
import net.thecir.reportmanagers.TechnopolisReportManager;

/**
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public class LiteReportManager {

    private final ExecutorService newFileExec = Executors.newFixedThreadPool(1);
    private final ExecutorService reportGeneratorExec = Executors.newFixedThreadPool(1);

    private static LiteReportManager SINGLETON;

    private JFrame parentFrame;
    private ReportManager reportManager;

    public static LiteReportManager getInstance() {
        if (SINGLETON == null) {
            SINGLETON = new LiteReportManager();
        }
        return SINGLETON;
    }

    public void initOutputComponents(JFrame parentFrame) {
        this.parentFrame = parentFrame;
    }

    public void createNewFile() {
        newFileExec.execute(() -> {
            try {
                File newFile = NewFileManager.getInstance().createNewWorkbook();
            } catch (OutputFileIOException ex) {
                log.log(Level.SEVERE, "An error occured while saving file.", ex);
                printMessageViaPane(ex.getMessage(), JOptionPane.ERROR_MESSAGE);
            } catch (NewFileCreationException ex) {
                log.log(Level.SEVERE, "An error occured while creating file.", ex);
                printMessageViaPane(ex.getMessage(), JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    public void generateReport(File inputFile, File outputFile, boolean undo, Stores store) {
        reportGeneratorExec.execute(() -> {
            if (store == Stores.Technopolis) {
                reportManager = new TechnopolisReportManager(inputFile, outputFile, undo);
            } else if (store == Stores.Technomarket) {
                reportManager = new TechnomarketReportManager(inputFile, outputFile, undo);
            }
            try {
                reportManager.generateReport();
            } catch (OutputFileIsFullException | OutputFileNoRecordsFoundException | InputFileNotMatchingSelectedFileException | OutputFileNotCorrectException | OutputFileIOException | InputFileContainsNoValidDateException ex) {
                log.log(Level.SEVERE, ex.getMessage(), ex);
                printMessageViaPane(ex.getMessage(), JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    private void printMessageViaPane(String message, int errorMessage) {
        SwingUtilities.invokeLater(() -> {
            ResourceBundle rb = ResourceBundle.getBundle("LanguageBundles/Bundle");
            JOptionPane.showMessageDialog(parentFrame, message, rb.getString("MessageTitle"), errorMessage);
        });
    }
}
