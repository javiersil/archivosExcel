/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author cosma_000
 */
public class ArchivoExcel {

    public Map<String, Object> leerArchivo(String urlFile) {
        InputStream excelStream = null;
        Map<String, Object> responseMap = new HashMap<String, Object>();
        try {

            excelStream = new FileInputStream(new File(urlFile));

             XSSFWorkbook hssfWorkbook = new  XSSFWorkbook(excelStream);

            XSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
            XSSFRow hssfRow;
            // HSSFCell cell;
            int rows = hssfSheet.getLastRowNum();

      

            for (int r = 0; r < rows; r++) {

                hssfRow = hssfSheet.getRow(r);

                if (hssfRow == null) {
                    break;
                } else {

                    String cellValue;

                    for (int c = 0; c <  hssfRow.getLastCellNum(); c++) {
                        
                        //......
                        cellValue = value(hssfRow.getCell(c)).trim();
                        System.out.print(" "+ cellValue);
                    }
                    System.out.println();
                }
            }
        } catch (FileNotFoundException fileNotFoundException) {

            System.out.println("El archivo no existe : " + fileNotFoundException.getMessage());

        } catch (IOException ex) {

            System.out.println("Error al procesar el archivo: " + ex.getMessage());

        } finally {

            try {
                excelStream.close();
            } catch (IOException ex) {

                System.out.println("Error en el proceso del archivo : " + ex.getMessage());

            }

        }
      return responseMap;
    }

    public String value(XSSFCell cell) {
        return cell == null ? ""
                : (cell.getCellType() == Cell.CELL_TYPE_STRING) ? cell.getStringCellValue()
                : (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) ? "" + cell.getNumericCellValue()
                : (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) ? "" + cell.getBooleanCellValue()
                : (cell.getCellType() == Cell.CELL_TYPE_BLANK) ? ""
                : (cell.getCellType() == Cell.CELL_TYPE_FORMULA) ? ""
                : (cell.getCellType() == Cell.CELL_TYPE_ERROR) ? "" : "";
    }

}
