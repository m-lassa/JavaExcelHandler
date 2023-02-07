module com.excelmetadata.excelpasswordremover {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;


    opens com.excelmetadata.excelpasswordremover to javafx.fxml;
    exports com.excelmetadata.excelpasswordremover;
}