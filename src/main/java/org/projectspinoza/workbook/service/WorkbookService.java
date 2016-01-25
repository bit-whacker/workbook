package org.projectspinoza.workbook.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookService implements Service {
    public static final String workbookPath = "workbook.xlsx";
    
    public void read() {
        // TODO Auto-generated method stub
        System.out.println("un implemented method!!!");
    }

    public void write() {
        // TODO Auto-generated method stub
        System.out.println("un implemented method!!!");
    }

    public void update(int sheet, int row, int column, double value) throws Exception {
        // TODO Auto-generated method stub
        
        FileInputStream fis = new FileInputStream(new File(workbookPath));
        XSSFWorkbook workbook = new XSSFWorkbook (fis);
        XSSFSheet xlsheet = workbook.getSheetAt(sheet);
        XSSFRow xlrow = xlsheet.getRow(row);
        XSSFCell xlcell = xlrow.getCell(column);
        
        xlcell.setCellValue(value);
        
        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        
        fis.close();
        FileOutputStream fos =new FileOutputStream(new File(workbookPath));
        workbook.write(fos);
        fos.close();
        
        System.out.println("Done");
    }

    public void delete() {
        // TODO Auto-generated method stub
        
    }

}
