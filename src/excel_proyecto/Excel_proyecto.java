/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel_proyecto;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 *
 * @author hazae
 */
public class Excel_proyecto {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args){
        // TODO code application logic here
      
      Workbook wb = new HSSFWorkbook();
        try ( OutputStream fileOut = new FileOutputStream("miarchivo.xls")) {
            Sheet sheet1 = wb.createSheet("Primer Hoja");
            Sheet sheet2 = wb.createSheet("Segunda Hoja");
            Sheet sheet = wb.createSheet("Tercer Hoja");
            Row row = sheet.createRow(0); //crear fila. se establece el indice a 0 hasta N                           
            row.createCell(0).setCellValue("Num"); // Columna A  
            row.createCell(1).setCellValue("Nombre"); // Columna B   
            row.createCell(2).setCellValue("Edad");// Columna C  
            row.createCell(3).setCellValue("Correo"); // Columna D 

            row = sheet.createRow(1); //crear fila 2
            row.createCell(0).setCellValue(1); // Columna A  
            row.createCell(1).setCellValue("Hazael Galindo"); // Columna B   
            row.createCell(2).setCellValue(25);// Columna C  
            row.createCell(3).setCellValue("haza.el78@hotmail.com"); // Columna D 
            
            row = sheet.createRow(2); //crear fila 3
            row.createCell(0).setCellValue(2); // Columna A  
            row.createCell(1).setCellValue("Martha Galindo"); // Columna B   
            row.createCell(2).setCellValue(55);// Columna C  
            row.createCell(3).setCellValue("betty29671@hotmail.com"); // Columna D
            
            row = sheet.createRow(3); //crear fila 2
            row.createCell(0).setCellValue(3); // Columna A  
            row.createCell(1).setCellValue("Ricardo Cortes"); // Columna B   
            row.createCell(2).setCellValue(52);// Columna C  
            row.createCell(3).setCellValue("ricardo_2@hotmail.com"); // Columna D
            
            
            wb.write(fileOut);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        }
    
    
    
   
}
