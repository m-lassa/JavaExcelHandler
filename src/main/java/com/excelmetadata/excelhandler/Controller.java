package com.excelmetadata.excelhandler;

import com.excelmetadata.excelhandler.datamodel.ExcelPasswordRemover;
import com.excelmetadata.excelhandler.datamodel.FileBackupHandler;
import com.excelmetadata.excelhandler.datamodel.MetadataHandler;
import javafx.fxml.FXML;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
     * It calls the backupFile method of the FileBackupHandler class.
     * Alerts are displayed to the user in case of successful/unsuccessful operations.
     */
    @FXML
    private void backUpSelected(){
        String filePath = filePathTextField.getText();
        String fileName = fileNameTextField.getText();
        String fullFilePath = filePath + File.separator + fileName;
        boolean isBackedUp;

        if (filePath.trim().equalsIgnoreCase("") || fileName.trim().equalsIgnoreCase("")){
            Alert emptyFieldAlert = new Alert(Alert.AlertType.INFORMATION);
            emptyFieldAlert.setTitle("No File Path and/or File Name");
            emptyFieldAlert.setHeaderText(null);
            emptyFieldAlert.setContentText("Please insert a valid file path and name.");
            emptyFieldAlert.showAndWait();
        } else {
            isBackedUp = FileBackupHandler.backupFile(fullFilePath);
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
     * It calls the removeExcelPassword method of the ExcelPasswordRemover class.
     * Alerts are displayed to the user in case of successful/unsuccessful operations.
     */
    @FXML
    private void removePassword(){
        String filePath = filePathTextField.getText();
        String fileName = fileNameTextField.getText();
        String fullFilePath = filePath + File.separator + fileName;
        boolean passRemoved;
        try{
            passRemoved = ExcelPasswordRemover.removeExcelPassword(fullFilePath);
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


    /**
     * Method called when the user clicks on the Display Excel Metadata button.
     * It creates an instance of the MetadataHandler class and calls its displayProperties method.
     * The metadata of the selected Excel file is dispplayed on a separate window.
     * Alerts are displayed to the user in case of successful/unsuccessful operations.
     */
    @FXML
    private void viewMetadata(){
        String filePath = filePathTextField.getText();
        String fileName = fileNameTextField.getText();
        String fullFilePath = filePath + File.separator + fileName;

        if (fullFilePath != null) {
            try (FileInputStream inputStream = new FileInputStream(fullFilePath)) {
                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                StringBuilder properties = new StringBuilder();

                properties.append("\n\n");
                MetadataHandler metadata = new MetadataHandler();
                properties.append(metadata.displayProperties(fullFilePath));

                Label propertiesLabel = new Label(properties.toString());

                VBox root = new VBox(propertiesLabel);
                Scene scene = new Scene(root, 450, 400);
                Stage stage = new Stage();
                stage.setScene(scene);
                stage.setTitle("Document Properties");
                stage.show();
            } catch (FileNotFoundException e){
                Alert emptyFieldAlert = new Alert(Alert.AlertType.INFORMATION);
                emptyFieldAlert.setTitle("No File Path and/or File Name");
                emptyFieldAlert.setHeaderText(null);
                emptyFieldAlert.setContentText("Please insert a valid file path and name.");
                emptyFieldAlert.showAndWait();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            Alert emptyFieldAlert = new Alert(Alert.AlertType.INFORMATION);
            emptyFieldAlert.setTitle("No File Path and/or File Name");
            emptyFieldAlert.setHeaderText(null);
            emptyFieldAlert.setContentText("Please insert a valid file path and name.");
            emptyFieldAlert.showAndWait();
        }
    }

}