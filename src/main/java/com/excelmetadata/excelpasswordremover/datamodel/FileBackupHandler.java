package com.excelmetadata.excelpasswordremover.datamodel;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

/**
 * Class created to handle file backup operations. It contains two methods:
 * One private method to validate that a file path is correct, and a public one
 * that contains the code to handle the backup itself.
 */
public class FileBackupHandler {


    /**
     * Creates a backup file in the same directory of the specified file.
     * "_backup" is added to the name of the file, before the file extension.
     * The method first checks the validity of the given file path, and it also
     * checks for IOExceptions.
     * @param filePath  The full file path as a String
     * @return          true if the backup operation was successful, false otherwise.
     */
    public boolean backupFile(String filePath) {
        if(!isValidFilePath(filePath)) {
            System.out.println("Please input a correct file path.");
            return false;
        } else {
            try {
                // Get the file name and directory of the original file
                File originalFile = new File(filePath);
                String fileName = originalFile.getName();
                String fileDirectory = originalFile.getParent();

                // Create the backup file name by adding "_backup" before the extension
                int extensionIndex = fileName.lastIndexOf(".");
                String backupFileName = fileName.substring(0, extensionIndex) + "_backup" + fileName.substring(extensionIndex);
                String backupFilePath = fileDirectory + File.separator + backupFileName;

                // Copy the original file to a backup file
                Files.copy(Paths.get(filePath), Paths.get(backupFilePath), StandardCopyOption.REPLACE_EXISTING);

                System.out.println("The original excel file has been backed up as: " + backupFilePath);
                return true;
            }  catch (IOException e){
                System.out.println("IOException: " + e.getMessage());
                e.printStackTrace();
                return false;
            }
        }
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
