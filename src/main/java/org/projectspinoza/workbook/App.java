package org.projectspinoza.workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.projectspinoza.workbook.service.WorkbookService;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "updating workbook" );
        try{
            WorkbookService workbook = new WorkbookService();
            workbook.update(0, 20, 2, 0.1);
        }catch(Exception ex){
            System.out.println("got exception at Main");
            ex.printStackTrace();
        }
    }
}
