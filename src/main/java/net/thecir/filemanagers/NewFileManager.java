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
package net.thecir.filemanagers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.FutureTask;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.SwingUtilities;
import lombok.extern.java.Log;
import net.thecir.callbacks.FileCallback;
import net.thecir.exceptions.NewFileCreationException;
import net.thecir.exceptions.OutputFileIOException;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Creates a new xlsx file and formats it.
 *
 * @author Konstantin Tsanov <k.tsanov@gmail.com>
 */
@Log
public class NewFileManager {

    private volatile FileCallback fileCallback;

    private static NewFileManager instance;

    private NewFileManager() {
    }

    public static NewFileManager getInstance() {
        if (instance == null) {
            instance = new NewFileManager();
        }
        return instance;
    }

    public void setFileCallback(FileCallback fileCallback) {
        this.fileCallback = fileCallback;
    }

    public File createNewWorkbook() throws OutputFileIOException, NewFileCreationException {
        XSSFWorkbook wb;
        System.out.println();
        try {
            wb = new XSSFWorkbook(ClassLoader.getSystemResourceAsStream("excel/default.xlsx"));
            NewFileFormatter formatter = new NewFileFormatter(wb);
            formatter.formatWorkbook();
            FutureTask<File> getFileTask = new FutureTask<>(new Callable<File>() {
                @Override
                public File call() {
                    return fileCallback.getFile();
                }
            });
            try {
                SwingUtilities.invokeAndWait(getFileTask);
            } catch (InterruptedException ex) {
                log.log(Level.SEVERE, "A thread waiting for the user to select new file to be created has been interrupted.", ex);
                throw new RuntimeException("Failed to create new file.");
            } catch (InvocationTargetException ex) {
                log.log(Level.SEVERE, "Invocation of the runnable to obtain the new file name/path has failed.", ex);
                throw new RuntimeException("Failed to create new file.");
            }

            File file = null;
            try {
                file = getFileTask.get();
            } catch (InterruptedException ex) {
                log.log(Level.SEVERE, "The thread waiting for the user to select new file to be created has been interrupted.", ex);
                throw new RuntimeException("Failed to create new file.");
            } catch (ExecutionException ex) {
                log.log(Level.SEVERE, "Failed to obtain the new file name.", ex);
                throw new RuntimeException("Failed to create new file.");
            }
            if (file == null) {
                return null;
            }
            if (!file.renameTo(file)) {
                log.log(Level.SEVERE, "The selected output file is in use by another process/program.");
                throw new OutputFileIOException("A file with the same name already exists and is in use by another program. Please close the file before attempting to save.");
            }
            if (!FilenameUtils.getExtension(file.getName()).equalsIgnoreCase("xlsx")) {
                if ("".equals(FilenameUtils.getExtension(file.getAbsolutePath()))) {
                    file = new File(file.toString() + ".xlsx");
                } else {
                    file = new File(file.getParentFile(), FilenameUtils.getBaseName(file.getName()) + ".xlsx");
                }
            }
            try (FileOutputStream fileOut = new FileOutputStream(file)) {
                wb.write(fileOut);
                return file;
            } catch (IOException ex) {
                throw new OutputFileIOException("A problem occured while saving file!");
            }

        } catch (OutputFileIOException ex) {
            throw ex;
        } catch (IOException ex) {
            log.log(Level.SEVERE, "A problem occured while getting the default workbook...", ex);
            throw new NewFileCreationException("Cannot create new file!");
        }
    }
}
