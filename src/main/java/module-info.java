module com.excelmetadata.excelhandler {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.poi;


    opens com.excelmetadata.excelhandler to javafx.fxml;
    exports com.excelmetadata.excelhandler;
}