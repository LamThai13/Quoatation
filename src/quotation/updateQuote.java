/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quotation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Kevin
 */
public class updateQuote {
    public void update(QuotationEntities quote,String path) throws FileNotFoundException, IOException{


       FileInputStream input = new FileInputStream(new File("C://Users//Kevin//Desktop//MS-ETN-quotation-Sample_.xlsx"));
       //create workbook;
       XSSFWorkbook wb = new XSSFWorkbook(input);
       //get sheet from workbook
       XSSFSheet sheet = wb.getSheetAt(0);
       //write out SN
          
       Row row0 = sheet.createRow(0);
       Cell r0c0 = row0.createCell(0);
       r0c0.setCellValue(quote.getSN());
       Row row2 = sheet.getRow(2);
       Cell r2c6 = row2.getCell(6);
       Cell r2c2 = row2.getCell(2);
       Cell r2c4 = row2.getCell(4);
       Cell r2c7 = row2.getCell(7);
       r2c7.setCellFormula("IF(B3*C3+D3*E3+F3*G3=0,\"\",B3*C3+D3*E3+F3*G3)");
       String Model = quote.getModel();
       switch(Model){
           case "M10":
               r2c2.setCellValue(1);
               break;
           case "M30":
               r2c4.setCellValue(1);
               break;
           case "M50":
               r2c6.setCellValue(1);
               break;
       }
       Row row6 = sheet.getRow(6);
       Cell r6c2 = row6.getCell(2);
       Cell r6c4 = row6.getCell(4);
       Cell r6c6 = row6.getCell(6);
       Cell r6c7 = row6.getCell(7);
       r6c7.setCellFormula("IF(B7*C7+D7*E7+F7*G7=0,\"\",B7*C7+D7*E7+F7*G7)");
       String top = quote.getTop();
       switch (Model){
           case "M10":               
            if(top.equalsIgnoreCase("y"))
            r6c2.setCellValue(1);
            break;
            case "M30":
            if(top.equalsIgnoreCase("y"))
            r6c4.setCellValue(1);
            break;
            case "M50":
            if(top.equalsIgnoreCase("y"))
            r6c6.setCellValue(1);
            break;    
    }
       Row row7 = sheet.getRow(7);
       Cell r7c2 = row7.getCell(2);
       Cell r7c4 = row7.getCell(4);
       Cell r7c6 = row7.getCell(6);
       Cell r7c7 = row7.getCell(7);
       r7c7.setCellFormula("IF(B8*C8+D8*E8+F8*G8=0,\"\",B8*C8+D8*E8+F8*G8)");
       String Bottom = quote.getBottom();
       switch (Model){
           case "M10":               
            if(Bottom.equalsIgnoreCase("y"))
            r7c2.setCellValue(1);
            break;
            case "M30":
            if(Bottom.equalsIgnoreCase("y"))
            r7c4.setCellValue(1);
            break;
            case "M50":
            if(Bottom.equalsIgnoreCase("y"))
            r7c6.setCellValue(1);
            break;    
    }
       Row row8 = sheet.getRow(8);
       Cell r8c6 = row8.getCell(6);
       Cell r8c7 = row8.getCell(7);
       r8c7.setCellFormula("IF(B9*C9+D9*E9+F9*G9=0,\"\",B9*C9+D9*E9+F9*G9)");
       if(quote.getModel().equalsIgnoreCase("M50")){
       if(quote.getFront().equalsIgnoreCase("y"))
           r8c6.setCellValue(1);
       }
     /*  Row row13 = sheet.getRow(13);
       Cell r13c2 = row13.getCell(2);
       if(list.get(1).getFront().equalsIgnoreCase("y"))
           r13c2.setCellValue(1);*/
       
       Row row14 = sheet.getRow(14);
       Cell r14c2 = row14.getCell(2);
       Cell r14c4 = row14.getCell(4);
       Cell r14c6 = row14.getCell(6);
       Cell r14c7 = row14.getCell(7);
       r14c7.setCellFormula("IF(B15*C15+D15*E15+F15*G15=0,\"\",B15*C15+D15*E15+F15*G15)");
       String powercord = quote.getPowercord();
       switch (Model){
           case "M10":               
            if(powercord.equalsIgnoreCase("n"))
            r14c2.setCellValue(1);
            break;
            case "M30":
            if(powercord.equalsIgnoreCase("n"))
            r14c4.setCellValue(1);
            break;
            case "M50":
            if(powercord.equalsIgnoreCase("n"))
            r14c6.setCellValue(1);
            break;    
    }
       Row row19 = sheet.getRow(19);
             
       Cell r19c2 = row19.getCell(2);
       Cell r19c4 = row19.getCell(4);
       Cell r19c6 = row19.getCell(6);
       Cell r19c7 = row19.getCell(7);
       r19c7.setCellFormula("IF(B20*C20+D20*E20+F20*G20=0,\"\",B20*C20+D20*E20+F20*G20)");
       String foam = quote.getFoam();
       switch (Model){
           case "M10":               
            if(foam.equalsIgnoreCase("n"))
            r19c2.setCellValue(1);
            break;
            case "M30":
            if(foam.equalsIgnoreCase("n"))
            r19c4.setCellValue(1);
            break;
            case "M50":
            if(foam.equalsIgnoreCase("n"))
            r19c6.setCellValue(1);
            break;    
    }
       Row row16 = sheet.getRow(16);
       Row row17 = sheet.getRow(17);       
       Cell r17c2 = row17.getCell(2);
       Cell r16c4 = row16.getCell(4);
       Cell r16c6 = row16.getCell(6);
       Cell r16c7 = row16.getCell(7);
       r16c7.setCellFormula("IF(B17*C17+D17*E17+F17*G17=0,\"\",B17*C17+D17*E17+F17*G17)");
       
       switch (Model){
           case "M10":               
            
            r17c2.setCellValue(1);
            break;
            case "M30":
            
            r16c4.setCellValue(1);
            break;
            case "M50":
            
            r16c6.setCellValue(1);
            break;    
    }
       String carton = quote.getBox();
       Row row18 = sheet.getRow(18);       
       
       Cell r18c4 = row18.getCell(4);
       Cell r18c6 = row18.getCell(6);
       Cell r18c7 = row18.getCell(7);
       r18c7.setCellFormula("IF(B19*C19+D19*E19+F19*G19=0,\"\",B19*C19+D19*E19+F19*G19)");
       
       switch (Model){
           case "M10":               
            if(carton.equalsIgnoreCase("n"))
            r17c2.setCellValue(1);
            break;
            case "M30":
            if(carton.equalsIgnoreCase("n"))
            r18c4.setCellValue(1);
            break;
            case "M50":
            if(carton.equalsIgnoreCase("n"))
            r18c6.setCellValue(1);
            break;    
    }
       Row row22 = sheet.getRow(22);       
       Cell r22c2 = row22.getCell(2);
       Cell r22c4 = row22.getCell(4);
       Cell r22c6 = row22.getCell(6);
       Cell r22c7 = row22.getCell(7);
       r22c7.setCellFormula("IF(B23*C23+D23*E23+F23*G23=0,\"\",B23*C23+D23*E23+F23*G23)");
       String accesory = quote.getDVI();
       switch (Model){
           case "M10":               
            if(accesory.equalsIgnoreCase("n"))
            r22c2.setCellValue(1);
            break;
            case "M30":
            if(accesory.equalsIgnoreCase("n"))
            r22c4.setCellValue(1);
            break;
            case "M50":
            if(accesory.equalsIgnoreCase("n"))
            r22c6.setCellValue(1);
            break;    
    }
       Row row47 = sheet.getRow(47);       
       Cell r47c2 = row47.getCell(2);
       Cell r47c4 = row47.getCell(4);
       Cell r47c6 = row47.getCell(6);
       Cell r47c7 = row47.getCell(7);
       r47c7.setCellFormula("IF(B48*C48+D48*E48+F48*G48=0,\"\",B48*C48+D48*E48+F48*G48)");
       String keyboard = quote.getKeyboard();
       switch (Model){
           case "M10":               
            if(keyboard.equalsIgnoreCase("n"))
            r47c2.setCellValue(1);
            break;
            case "M30":
            if(keyboard.equalsIgnoreCase("n"))
            r47c4.setCellValue(1);
            break;
            case "M50":
            if(keyboard.equalsIgnoreCase("n"))
            r47c6.setCellValue(1);
            break;    
    }
       Row row48 = sheet.getRow(48);       
       Cell r48c2 = row48.getCell(2);
       Cell r48c4 = row48.getCell(4);
       Cell r48c6 = row48.getCell(6);
       Cell r48c7 = row48.getCell(7);
       r48c7.setCellFormula("IF(B49*C49+D49*E49+F49*G49=0,\"\",B49*C49+D49*E49+F49*G49)");
       double portC = quote.getPortCover();
        System.out.println(portC);
       switch (Model){
           case "M10":               
            if(portC!=0.0)
            r48c2.setCellValue(portC);
            break;
            case "M30":
            if(portC!=0.0)
            r48c4.setCellValue(portC);
            break;
            case "M50":
            if(portC!=0.0)
            r48c6.setCellValue(portC);
            break;    
    }
       Row row49 = sheet.getRow(49);       
       Cell r49c2 = row49.getCell(2);
       Cell r49c4 = row49.getCell(4);
       Cell r49c6 = row49.getCell(6);
      
       
       switch (Model){
           case "M10":                          
            r49c2.setCellValue(1);
            break;
            case "M30":            
            r49c4.setCellValue(1);
            break;
            case "M50":            
            r49c6.setCellValue(1);
            break;    
    }
       sheet.getRow(49).getCell(7).setCellFormula("IF(B50*C50+D50*E50+F50*G50=0,\"\",B50*C50+D50*E50+F50*G50)");
       sheet.getRow(50).getCell(7).setCellFormula("SUM(H3:H50)");
       String fileName="MS-ETN-quotation-"+quote.getSN()+".xlsx";
       File f = new File("C://Users//Kevin//Desktop//MilestoneQuote//"+fileName);
       
       FileOutputStream fos =new FileOutputStream(new File(f.getAbsolutePath()));
	        wb.write(fos);
	        fos.close();
       
    }
}
