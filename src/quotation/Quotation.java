/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package quotation;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.Statement;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.table.TableColumnModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Quotation {


    public static void main(String[] args) throws IOException {
        /*inputform formFrame = new inputform();
        formFrame.setVisible(true);
        formFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);*/
        
        ReadExcelFileToList reader = new ReadExcelFileToList();
        List<QuotationEntities> list ;
        //String path ="C:/Users/Kevin/Desktop/visualizecheck.xlsx"; 
        String path =pathFile().toString();
        list = reader.readExcelData(path);
        //System.out.println(list);       
        updateQuote update = new updateQuote();
        Iterator<QuotationEntities> ListIter = list.iterator();
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Select location to save your Quotations");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.showSaveDialog(null);
        
        while(ListIter.hasNext()){
            QuotationEntities quote = ListIter.next();
            update.update(quote, path,chooser.getSelectedFile().getAbsolutePath());
        }
        
      /* Connection conn = null;
        Statement st = null;
        ResultSet rs = null; 
                    try {
            String url = "jdbc:mysql://localhost/receptioncheck";
            String username = "root";
            String password = "";
            conn = (Connection) DriverManager.getConnection(url,username,password);
                
                if (conn != null){
                    System.out.println("OK");
                }
                
                String sql = "select * from visualization";
                
                st = (Statement) conn.createStatement();
                
                rs = st.executeQuery(sql);
            /*    Vector data = null;
                
                tblModel.setRowCount(0);
                
                if (rs.isBeforeFirst()==false){
                JOptionPane.showMessageDialog(this, "No data");
                return;
                }
               
                while (rs.next()){
                data = new Vector();
                data.add(rs.getString("SN"));
                data.add(rs.getString("RMA"));
                data.add(rs.getString("Model"));
                data.add(rs.getBoolean("Box"));
                data.add(rs.getBoolean("Foam"));
                data.add(rs.getBoolean("PowerCord"));
                data.add(rs.getBoolean("DVI"));
                data.add(rs.getBoolean("Keyboard"));
                data.add(rs.getBoolean("VideoCable"));
                data.add(rs.getBoolean("Top"));
                data.add(rs.getBoolean("Bottom"));
                data.add(rs.getBoolean("Front"));
                data.add(rs.getBoolean("Warranty"));
                data.add(rs.getInt("PortCover"));
                
                tblModel.addRow(data);
                
                }
                jTable1.setModel(tblModel);
                TableColumnModel column = jTable1.getColumnModel();
                column.getColumn(0).setPreferredWidth(150);
                column.getColumn(1).setPreferredWidth(130);
                column.getColumn(2).setPreferredWidth(50);
            } catch (SQLException ex) {
                
            }  */
    }
    public static String pathFile(){
            
            JFileChooser chooser = new JFileChooser();
            String fileName="";
            chooser.setDialogTitle("Select visualize-check file");                       
            int result = chooser.showOpenDialog(null);
            if(result == JFileChooser.APPROVE_OPTION){
                File f = chooser.getSelectedFile();
                fileName = f.getAbsolutePath();
                chooser.setVisible(true);
        }
            return fileName;
    }
}
