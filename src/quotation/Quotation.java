/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quotation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFrame;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Quotation {


    public static void main(String[] args) throws IOException {
        inputform formFrame = new inputform();
        formFrame.setVisible(true);
        
        /*
        ReadExcelFileToList reader = new ReadExcelFileToList();
        List<QuotationEntities> list ;
        String path ="C:/Users/Kevin/Desktop/visualizecheck.xlsx"; 
        list = reader.readExcelData(path);       
        updateQuote update = new updateQuote();
        Iterator<QuotationEntities> ListIter = list.iterator();
        while(ListIter.hasNext()){
            QuotationEntities quote = ListIter.next();
            update.update(quote, path);
        }*/
    }
}
