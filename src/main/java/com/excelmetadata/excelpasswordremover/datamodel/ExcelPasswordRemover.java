package com.excelmetadata.excelpasswordremover.datamodel;

import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Class created to handle spreadsheet password removal operations. It contains two methods:
 * One private method to validate that a file path is correct, and a public one
 * that contains the code to remove workbook and worksheet protection from .xlsx/.xls files.
 */
public class ExcelPasswordRemover {

    /**
     * Removes workbook password protection, as well as worksheet password protection from
     * every sheet in the Excel file. The method currently handles both .xlsx and .xls
     * spreasheet formats.
     * @param filePath  The full file path as a String
     * @return          true if the operation was successful, false otherwise.
     * @throws IOException
     */
    public boolean removeExcelPassword(String filePath)
            throws IOException {

        if(!isValidFilePath(filePath)) {
            System.out.println("Please input a correct file path.");
            return false;
        }

        FileInputStream inputStream = new FileInputStream(new File(filePath));
        String extension = filePath.substring(filePath.lastIndexOf(".") + 1);

        if(extension.equalsIgnoreCase("xlsx") || extension.equalsIgnoreCase("xls")){
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

            // Remove password protection from all sheets in the workbook
            int numberOfSheets = workbook.getNumberOfSheets();
            if (numberOfSheets > 0) {
                // Remove password protection from all sheets in the workbook
                for (int i = 0; i < numberOfSheets; i++) {
                    if(workbook.getSheetAt(i).getProtect()) {
                        workbook.getSheetAt(i).protectSheet(null);
                    }

                }
                // Remove password protection from the workbook
                workbook.setWorkbookPassword(null, HashAlgorithm.sha512);

                // Save the workbook to the same file path
                FileOutputStream outputStream = new FileOutputStream(filePath);
                workbook.write(outputStream);
                outputStream.close();
                workbook.close();
                return true;
            }
        } else return false;
        return false;
    }

    /**
     * Returns a boolean value based on the existence of the specified file path.
     * It relies on the exists() method of the java.io.File class to check whether the file
     * at the specified directory location exists
     * @param filePath  The full file path as a String
     * @return          true if the file with the specified path exists, false if it doesn't
     */
    private boolean isValidFilePath(String filePath) {
        File file = new File(filePath);
        return file.exists() && file.isFile();
    }
}
