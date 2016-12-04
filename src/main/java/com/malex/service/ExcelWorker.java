package com.malex.service;

import com.malex.model.DataModel;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Random;

/**
 * Class {@link ExcelWorker} creates Excel files.
 *
 * @author malex
 */
public class ExcelWorker {

    /**
     * Create excel file.
     *
     * @param fileName      the name of the excel file
     * @param sheetName     the name of sheet in the excel file
     * @param dataModelList the data for the excel file
     */
    public static void createExcelFile(String fileName, String sheetName, List<DataModel> dataModelList) {

        // #1 Create the excel file in memory
        HSSFWorkbook workBook = new HSSFWorkbook();

        // #2 Creating a sheet called "Simple sheet"
        HSSFSheet sheet = workBook.createSheet(sheetName);

        // #3 Fill the list with some data
        List<DataModel> dataList = dataModelList;

        // #4 Counter for lines
        int rowNum = 0;

        // #5 Create captions for the columns
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue("First Name");
        row.createCell(1).setCellValue("Last Name");
        row.createCell(2).setCellValue("City");
        row.createCell(3).setCellValue("Salary");

        // #6 Fill the data in the sheet
        for (DataModel dataModel : dataList) {
            //increment row
            rowNum = rowNum + 1;
            //fill the sheet
            createSheetHeader(sheet, rowNum, dataModel);
        }

        // Recording the created Excel document in memory to a file
        try (FileOutputStream out = new FileOutputStream(new File(fileName + "_" + new Random().nextInt(50) + ".xls"))) {
            workBook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel file created successfully!");

    }

    /**
     * Filling line (rowNum) specified sheet (sheet).
     */
    private static void createSheetHeader(HSSFSheet sheet, int rowNum, DataModel dataModel) {
        Row row = sheet.createRow(rowNum);

        row.createCell(0).setCellValue(dataModel.getName());
        row.createCell(1).setCellValue(dataModel.getSurname());
        row.createCell(2).setCellValue(dataModel.getCity());
        row.createCell(3).setCellValue(dataModel.getSalary());
    }


}
