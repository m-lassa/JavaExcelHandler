package com.excelmetadata.excelpasswordremover;

import com.excelmetadata.excelpasswordremover.datamodel.ExcelPasswordRemover;
import com.excelmetadata.excelpasswordremover.datamodel.FileBackupHandler;
import javafx.fxml.FXML;
import javafx.scene.control.Alert;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;

import java.io.File;
import java.io.IOException;

public class Controller {

    @FXML
    private VBox vBox;
    @FXML
    private TextField filePathTextField;
    @FXML
    private TextField fileNameTextField;

    /**
     * Method called when the user clicks on the Back Up File button.
     * It creates an instance of the FileBackupHandler class and calls its backupFile method.
     * Alerts are displayed to the user in case of successful/unsuccessful operations.
     */
    @FXML
    private void backUpSelected(){
        String filePath = filePathTextField.getText();
        String fileName = fileNameTextField.getText();
        String fullFilePath = filePath + File.separator + fileName;
        FileBackupHandler backupHandler = new FileBackupHandler();
        boolean isBackedUp;

        if (filePath.trim().equals("") || fileName.trim().equalsIgnoreCase("")){
            Alert emptyFieldAlert = new Alert(Alert.AlertType.INFORMATION);
            emptyFieldAlert.setTitle("No File Path and/or File Name");
            emptyFieldAlert.setHeaderText(null);
            emptyFieldAlert.setContentText("Please insert a valid file path and name.");
            emptyFieldAlert.showAndWait();
        } else {
            isBackedUp = backupHandler.backupFile(fullFilePath);
            if(!isBackedUp){
                Alert wrongPathAlert = new Alert(Alert.AlertType.ERROR);
                wrongPathAlert.setContentText("Error backing up file. Directory or file do not exist");
                wrongPathAlert.setHeaderText(null);
                wrongPathAlert.showAndWait();
            } else if(isBackedUp){
                Alert correctPathInfoAlert = new Alert(Alert.AlertType.CONFIRMATION);
                correctPathInfoAlert.setContentText("Back up executed correctly");
                correctPathInfoAlert.setHeaderText(null);
                correctPathInfoAlert.showAndWait();
            }
        }
    }

    /**
     * Method called when the user clicks on the Choose File button.
     * It uses an instance of javafx.stage.FileChooser to allow the user
     * to select .xlsx or .xls files from the appropriate directory.
     * Directory name and file name are then used to populate the UI Text Fields.
     */
    @FXML
    private void openFile() {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Open Excel file");
        chooser.getExtensionFilters().add
                (new FileChooser.ExtensionFilter("Microsoft Excel", "*.xlsx", "*.xls"));

        File file = chooser.showOpenDialog(vBox.getScene().getWindow());

        if(file == null){
            return;
        }

        filePathTextField.setText(file.getParent());
        fileNameTextField.setText(file.getName());
    }

    /**
     * Method called when the user clicks on the Remove Spreadsheet Password button.
     * It creates an instance of the ExcelPasswordRemover class and calls its removeExcelPassword method.
     * Alerts are displayed to the user in case of successful/unsuccessful operations.
     */
    @FXML
    private void removePassword(){
        String filePath = filePathTextField.getText();
        String fileName = fileNameTextField.getText();
        String fullFilePath = filePath + File.separator + fileName;
        ExcelPasswordRemover passwordRemover = new ExcelPasswordRemover();
        boolean passRemoved;
        try{
            passRemoved = passwordRemover.removeExcelPassword(fullFilePath);
            if(!passRemoved){
                Alert incorrectRemovalAlert = new Alert(Alert.AlertType.ERROR);
                incorrectRemovalAlert.setContentText("Error removing workbook or worksheet password. \nEnsure the file has .xlsx or .xls as extension");
                incorrectRemovalAlert.setHeaderText(null);
                incorrectRemovalAlert.showAndWait();
            } else if(passRemoved){
                Alert correctRemovalAlert = new Alert(Alert.AlertType.CONFIRMATION);
                correctRemovalAlert.setContentText("Password successfully removed from workbook and/or worksheet \nIf no password was set, no change was made.");
                correctRemovalAlert.setHeaderText(null);
                correctRemovalAlert.showAndWait();
            }
        } catch (IOException e){
            e.printStackTrace();
        }
    }

}