/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class OfficeUtils {
    
    /**
     * Con este metodo leeremos los archivos en modo de compatibilidad, es decir los archivos
     * con extencion .xls o creados con Microsoft Office 2007 o anteriores.
     * 
     * @param Excel2003FileToRead ruta del archivo con terminacion xls, si null es recivido una excepcion es arrojada.
     * @throws java.io.IOException 
     */
    public static void readXLSFile(FileInputStream Excel2003FileToRead) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(Excel2003FileToRead);
        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        Iterator rows = sheet.rowIterator();
        while (rows.hasNext()) {
            row = (HSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            while (cells.hasNext()) {
                cell = (HSSFCell) cells.next();
                if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
                    System.out.print(cell.getStringCellValue() + " ");
                } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(cell.getNumericCellValue() + " ");
                } else {
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println();
        }
    }

    /**
     * Con este metodo leeremos los archivos en modo de compatibilidad, es decir los archivos
     * con extencion .xlsx o creados con Microsoft Office 2010 o posteriores.
     * Algoritmo y funcionamiento de la libreria POI-3.10 para lectura de Excel:
     *      Al generar un libro, le entregamos la ruta de el archivo con formato xlsx
     *      una vez teniendo la instancia de el libro seleccionamos la hoja con la cual trabajar.
     *      es decir (XSSSheet = 0 es la primera hoja), una vez seleccionada la hoja usamos las
     *      estructuras de la libreria para obtener las filas, o celdas individualmente.
     * 
     * @param Excel2007FileToRead
     * @throws IOException 
     */
    public static void readXLSXFile(FileInputStream Excel2007FileToRead) throws IOException{
        XSSFWorkbook xwb = new XSSFWorkbook(Excel2007FileToRead);
        XSSFSheet xsheet = xwb.getSheetAt(0);
        XSSFRow xrow;
        XSSFCell xcell;
        Iterator xrows = xsheet.rowIterator();
        while (xrows.hasNext()) {
            xrow = (XSSFRow) xrows.next();
            Iterator xcells = xrow.cellIterator();
            while (xcells.hasNext()) {
                xcell = (XSSFCell) xcells.next();
                if (xcell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                    System.out.print(xcell.getStringCellValue() + " ");
                } else if (xcell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                    System.out.print(xcell.getNumericCellValue() + " ");
                } else {
                    //U Can Handel Boolean, Formula, Errors
                }
            }
            System.out.println("");
        }
    }
    
       public static void main(String[] args) {
        try {
            File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Plantillas_salida_Listasprecios.xlsx");
            String fileToReadname = fileToRead.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String excel2003 = "xls";
            String excel2007 = "xlsx";
            if (excel2003.equalsIgnoreCase(extension)) {
                FileInputStream Excel2003FileToRead = new FileInputStream(fileToRead);
                readXLSFile(Excel2003FileToRead);
            } else if (excel2007.equalsIgnoreCase(extension)) {
                FileInputStream Excel2007FileToRead = new FileInputStream(fileToRead);
                readXLSXFile(Excel2007FileToRead);
            }
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
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
