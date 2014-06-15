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
    
    
    public void readPDF(){
        PDDocument pd;
        BufferedWriter wr;
        
        try {
         File input = new File("C:\\Invoice.pdf");  // The PDF file from where you would like to extract
         File output = new File("C:\\SampleText.txt"); // The text file where you are going to store the extracted data
         pd = PDDocument.load(input);
         System.out.println(pd.getNumberOfPages());
         System.out.println(pd.isEncrypted());
         pd.save("CopyOfInvoice.pdf"); // Creates a copy called "CopyOfInvoice.pdf"
         PDFTextStripper stripper = new PDFTextStripper();
         stripper.setStartPage(3); //Start extracting from page 3
         stripper.setEndPage(5); //Extract till page 5
         wr = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output)));
         stripper.writeText(pd, wr);
         if (pd != null) {
             pd.close();
         }
         
        // I use close() to flush the stream.
        wr.close();
        } catch (Exception e){
         e.printStackTrace();
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
    
}
