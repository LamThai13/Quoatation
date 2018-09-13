/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quotation;

import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Kevin
 */
public class WriteJframeatoExcelFile {
    public void writeJframetoFile(String fileName, List<QuotationEntities> Listquote, int rowIndex) throws Exception{
        Workbook workbook = null;
        if(fileName.toLowerCase().endsWith("xlsx")||fileName.toLowerCase().endsWith("xlsm")){
            workbook = new XSSFWorkbook();
        }else if(fileName.toLowerCase().endsWith("xls")){
            workbook = new HSSFWorkbook();
        }
        else throw new Exception("invalid file name, should be excel file");
        
        //create sheet
        Sheet sheet = workbook.createSheet("sheetName");
        Iterator<QuotationEntities> IterQuote = Listquote.iterator();
        Row firstRow = sheet.createRow(0);
        writeFirstline(firstRow); 
        while(IterQuote.hasNext()){
            QuotationEntities quote = IterQuote.next();
            Row row = sheet.createRow(rowIndex++);             
               writeCell(quote,row);
        }   
        
        FileOutputStream outputstream = new FileOutputStream(fileName);
        workbook.write(outputstream);
    }
    private void writeCell(QuotationEntities quote,Row row){
        Cell cell = row.createCell(0);
        cell.setCellValue(quote.getSN());
        
        cell = row.createCell(1);
        cell.setCellValue(quote.getRMA());
        
        /*cell = row.createCell(2);
        cell.setCellValue(quote.get);*/
        
        cell = row.createCell(3);
        cell.setCellValue(quote.getPartNo());
        
        cell = row.createCell(4);
        cell.setCellValue(quote.getDate());
        
        cell = row.createCell(5);
        cell.setCellValue(quote.getModel());
        
           
        cell = row.createCell(6);
        cell.setCellValue(quote.getBox());
        
        cell = row.createCell(7);
        cell.setCellValue(quote.getFoam());
        
        cell = row.createCell(8);
        cell.setCellValue(quote.getPowercord());
        
        cell = row.createCell(9);
        cell.setCellValue(quote.getDVI());
        
        cell = row.createCell(10);
        cell.setCellValue(quote.getKeyboard());
        
        cell = row.createCell(11);
        cell.setCellValue(quote.getVideocable());
        
        cell = row.createCell(12);
        cell.setCellValue(quote.getTop());
        
        cell = row.createCell(13);
        cell.setCellValue(quote.getBottom());
        
        cell = row.createCell(14);
        cell.setCellValue(quote.getFront());
        
        cell = row.createCell(15);
        cell.setCellValue(quote.getWarranty());
        
        cell = row.createCell(16);
        cell.setCellValue(quote.getPhoto());
                    
        cell = row.createCell(17);
        cell.setCellValue(quote.getPortCover());
        
        cell = row.createCell(18);
        cell.setCellValue(quote.getBootUpPass());
        
        cell = row.createCell(19);
        cell.setCellValue(quote.getBootUpPass());
    }
        public void writeFirstline(Row row){
        Cell cell = row.createCell(0);
        cell.setCellValue("SN");
        
        cell = row.createCell(1);
        cell.setCellValue("RMA #");
        
        /*cell = row.createCell(2);
        cell.setCellValue(quote.get);*/
        
        cell = row.createCell(3);
        cell.setCellValue("Part #");
        
        cell = row.createCell(4);
        cell.setCellValue("Date");
        
        cell = row.createCell(5);
        cell.setCellValue("Model");
        
           
        cell = row.createCell(6);
        cell.setCellValue("Box");
        
        cell = row.createCell(7);
        cell.setCellValue("Foam");
        
        cell = row.createCell(8);
        cell.setCellValue("Power cord");
        
        cell = row.createCell(9);
        cell.setCellValue("DVI");
        
        cell = row.createCell(10);
        cell.setCellValue("Keyboard&Mouse");
        
        cell = row.createCell(11);
        cell.setCellValue("Video Cable");
        
        cell = row.createCell(12);
        cell.setCellValue("Top");
        
        cell = row.createCell(13);
        cell.setCellValue("Bottom");
        
        cell = row.createCell(14);
        cell.setCellValue("Font");
        
        cell = row.createCell(15);
        cell.setCellValue("Warranty label");
        
        cell = row.createCell(16);
        cell.setCellValue("Photo");
                    
        cell = row.createCell(17);
        cell.setCellValue("Port cover");
        
        cell = row.createCell(18);
        cell.setCellValue("Boot up pass");
        
        cell = row.createCell(19);
        cell.setCellValue("Boot up fail");
    }
    
    /*private void writeCell(POSreport pos,Row row){
        Cell cell = row.createCell(0);       
           cell.setCellValue(pos.getSN()+";"+pos.getCompany()+";"+pos.getModel()+";"+pos.getDMIrev()+";"+pos.getDMIsn());
        
            
    }*/
}
