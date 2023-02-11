package com.excelmetadata.excelhandler.datamodel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * This class contains a method to extract metadata from an excel file.
 */
public class MetadataHandler {
    private String author;
    private String title;
    private String subject;
    private String keywords;
    private String description;
    private String numOfSheets;
    private String lastModifiedBy;
    private String created;
    private String modified;
    private String version;
    private String contentStatus;
    private String applicatiom;
    private String appVersion;
    private String company;

    /**
     * The method extracts different categories of metadata using the Apache POI
     * to create a XSSFWorkbook instance from the specified file path.
     * @param filePath
     * @return      The formatted list of workbook properties as a String
     */
    public String displayProperties(String filePath){
    try (FileInputStream inputStream = new FileInputStream(new File(filePath))){
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        StringBuilder sb = new StringBuilder();

        author = "Author: " + workbook.getProperties().getCoreProperties().getCreator() + "\n";
        title = "Title: " +  workbook.getProperties().getCoreProperties().getTitle() + "\n";
        subject = "Subject: " + workbook.getProperties().getCoreProperties().getSubject() + "\n";
        keywords = "Keywords: " + workbook.getProperties().getCoreProperties().getKeywords() + "\n";
        numOfSheets = "Number of sheets: " + workbook.getNumberOfSheets() + "\n";
        description = "Description: " + workbook.getProperties().getCoreProperties().getDescription() + "\n";
        lastModifiedBy = "Last Modified By: " + workbook.getProperties().getCoreProperties().getLastModifiedByUser() + "\n";
        created = "Created: " + workbook.getProperties().getCoreProperties().getCreated() + "\n";
        modified = "Modified: " + workbook.getProperties().getCoreProperties().getModified() + "\n";
        version = "Version: " + workbook.getProperties().getCoreProperties().getVersion() + "\n";
        contentStatus = "Content status: " +  workbook.getProperties().getCoreProperties().getContentStatus() + "\n";
        applicatiom= "Application: " + workbook.getProperties().getExtendedProperties().getApplication() + "\n";
        appVersion = "Application version: " + workbook.getProperties().getExtendedProperties().getAppVersion() + "\n";
        company = "Company: " + workbook.getProperties().getExtendedProperties().getCompany() + "\n";

        sb.append(author);
        sb.append(title);
        sb.append(subject);
        sb.append(keywords);
        sb.append(description);
        sb.append(numOfSheets);
        sb.append(lastModifiedBy);
        sb.append(created);
        sb.append(modified);
        sb.append(version);
        sb.append(contentStatus);
        sb.append(applicatiom);
        sb.append(appVersion);
        sb.append(company);
        sb.append("\n");

        if(workbook.getNumberOfSheets() > 1){
            for(int i = 0; i < workbook.getNumberOfSheets(); i++){
                sb.append("Sheet " + (i+1) + " info: " + workbook.getSheetAt(i) + "\n");
            }
        }

        return sb.toString();

    } catch (IOException e){
        e.printStackTrace();
        return null;
    }
}
}
