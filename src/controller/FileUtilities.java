/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

//JavaIO     
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;

//Collections
import java.util.ArrayList;
import java.util.Iterator;
// Actualice a ArrayList ya que tecnicamente Vector es obsoleto desde jdk 1.2
//import java.util.Vector; 

//Librerias para Apache PDFBox
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;

/* Librerias exclusivas de Apache POI */
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Utilerias:
 *  En esta clase se localizan las funciones necesarias para la obtencion y
 *  clasificacion de diversos datos, ademas de su uso en la aplicacion se 
 *  plantean metodos para la administracion basica, buscar, borrar, agregar,
 *  unir, crear, separar, etc.
 *  
 * La funcionalidad basica recibe la ruta del archivo a trabajar, verificar que exista,
 * asi como identificar el tipo de archivo (PDF, XLSX, DOCX), para despues trabajar con
 * el mismo.
 * 
 * Casos de uso:
 * 1.- FileUtilities fu = new FileUtilities("ruta/nombre.extencion");
 *  int numHojas = fu.obtenerHojas();
 * 
 * 2.- Iterator column = getColumns(new FileUtilities("ruta/nombre.extencion"));
 * 
 * 
 * @author Isaac Alcocer.
 */
public class FileUtilities {

    public String path;
    
    public FileUtilities(String path){
        this.path = path;
    }
    
    /*
    public static void main(String[] args) throws Exception {
		String filename = "\\Users\\Manuu Alcocer\\Downloads\\salarios.xlsx";
		FileInputStream fis = null;

		try {

			fis = new FileInputStream(filename);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator rowIter = sheet.rowIterator(); 

			while(rowIter.hasNext()){
				XSSFRow myRow = (XSSFRow) rowIter.next();
				Iterator cellIter = myRow.cellIterator();
				ArrayList<String> cellStoreVector=new ArrayList<>();
				while(cellIter.hasNext()){
					XSSFCell myCell = (XSSFCell) cellIter.next();
					String cellvalue = myCell.getRawValue();
					cellStoreVector.add(cellvalue);
				}
				String firstcolumnValue = null;
				String secondcolumnValue = null;

				int i = 0;
				firstcolumnValue = cellStoreVector.get(i); 
                                try {
				secondcolumnValue = cellStoreVector.get(i+1);
                                } catch (IndexOutOfBoundsException IOB) {System.out.println(IOB.toString());}
				insertQuery(firstcolumnValue,secondcolumnValue);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fis != null) {
				fis.close();
			}
		}

//		showExelData(sheetData);
	}

    private static void insertQuery(String firstcolumnvalue,String secondcolumnvalue) {
		System.out.println(firstcolumnvalue +  " "  +secondcolumnvalue);
    }
*/ 
    
}//fin
