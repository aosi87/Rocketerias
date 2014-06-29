/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import javax.swing.JFrame;
import javax.swing.ProgressMonitor;
import javax.swing.ProgressMonitorInputStream;
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
    
    private int numSheetTabs = 0;
    private int numRows = 0;
    private int numColumns = 0;
    private String nameSheetTab = "";
    //private ProgressMonitorInputStream Excel2003FileToRead;
    //private ProgressMonitorInputStream Excel2007FileToRead;
    
    private HSSFWorkbook xlsWorkbook = null; //Office 2003 (xls)
    private HSSFSheet sheetXLS = null;
    private XSSFWorkbook xlsxWorkbook = null; //Office 2007 (xlsx)
    private XSSFSheet sheetXLSX = null;
    
    public OfficeUtils(){}
    
    public OfficeUtils(HSSFWorkbook xlsFile){
        this.xlsWorkbook = xlsFile;
    }
    
    public OfficeUtils(XSSFWorkbook xlsxFile){
        this.xlsxWorkbook = xlsxFile;
    }
    
//    public OfficeUtils(File file, JFrame frame){
//            try {
//            String fileToReadname = file.getName();
//            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
//                    + 1, fileToReadname.length());
//            String excel2003 = "xls";
//            String excel2007 = "xlsx";
//            if (excel2003.equalsIgnoreCase(extension)) {
//                Excel2003FileToRead = new ProgressMonitorInputStream(frame,"Leyendo",new FileInputStream(file));
//                Excel2003FileToRead = new FileInputStream(file);
//                this.xlsWorkbook = new HSSFWorkbook(Excel2003FileToRead);
//                readXLSFile(Excel2003FileToRead);
//              } else if (excel2007.equalsIgnoreCase(extension)) {
//                Excel2007FileToRead = new ProgressMonitorInputStream(frame,"Leyendo",new FileInputStream(file));  
//                Excel2007FileToRead = new FileInputStream(file);
//                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToRead);
//                readXLSXFile(Excel2007FileToRead);
//              }
//            } catch (IOException ex) {
//                System.out.println(ex.getMessage());
//              }
//    }
    
    public OfficeUtils(File file){
            try {
            String fileToReadname = file.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String excel2003 = "xls";
            String excel2007 = "xlsx";
            if (excel2003.equalsIgnoreCase(extension)) {
                FileInputStream Excel2003FileToRead = new FileInputStream(file);
                this.xlsWorkbook = new HSSFWorkbook(Excel2003FileToRead);
                //readXLSFile(Excel2003FileToRead);
              } else if (excel2007.equalsIgnoreCase(extension)) {
                FileInputStream Excel2007FileToRead = new FileInputStream(file);
                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToRead);
                //readXLSXFile(Excel2007FileToRead);
              }
            } catch (IOException ex) {
                System.out.println(ex.getMessage());
              }
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
    
    public void readExcel(){
        if(this.xlsWorkbook==null){
        XSSFSheet xsheet = this.xlsxWorkbook.getSheetAt(0);
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
        } else {
            HSSFSheet sheet = this.xlsWorkbook.getSheetAt(0);
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
            System.out.print("");
        }
        
        }
    }
    
      public ArrayList createArrayList(){
        ArrayList filas = new ArrayList<>();
        ArrayList celdas  = new ArrayList<>();
        int numColumn=0;
        int numRow = 0;
        XSSFSheet xsheet = this.xlsxWorkbook.getSheetAt(0);
        XSSFRow xrow;
        XSSFCell xcell;
        Iterator xrows = xsheet.rowIterator();
        while (xrows.hasNext()) {
            //filas.add(celdas);
            xrow = (XSSFRow) xrows.next();
            numRow = xrow.getRowNum()+1;
            Iterator xcells = xrow.cellIterator();
            
            while (xcells.hasNext()) {
                xcell = (XSSFCell) xcells.next();
                numColumn = xcell.getColumnIndex()+1;
                
                switch(xcell.getCellType()){
                    case XSSFCell.CELL_TYPE_BLANK:
                        celdas.add("");
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        celdas.add(xcell.getNumericCellValue());
                        //System.out.print(xcell.getNumericCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_STRING:
                        celdas.add(xcell.getStringCellValue());
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    default:
                        //if(xcell.getRowIndex() == numColumn)  
                        //System.out.print("");
                        break;
                }
            }
        }
        
        System.out.println("filas: "+numRow+" columnas: "+numColumn);
        print(celdas);
        return filas;
    }
      
    public void print(ArrayList a){
        for (Object a1 : a) 
           // for (Object b1: (ArrayList)a1) 
                System.out.println( a1.toString() );   
    }

    public static void main(String[] args) {
        File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\plantilla salida.xlsx");
        //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Plantillas_salida_Listasprecios.xlsx");
        //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Date test.xlsx"); 
        OfficeUtils ou = new OfficeUtils(fileToRead);
        ou.createArrayList();
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

    /**
     * @return the numSheetTabs
     */
    public int getNumSheetTabs() {
        return numSheetTabs;
    }

    /**
     * @param numSheetTabs the numSheetTabs to set
     */
    public void setNumSheetTabs(int numSheetTabs) {
        this.numSheetTabs = numSheetTabs;
    }

    /**
     * @return the numRows
     */
    public int getNumRows() {
        return numRows;
    }

    /**
     * @param numRows the numRows to set
     */
    public void setNumRows(int numRows) {
        this.numRows = numRows;
    }

    /**
     * @return the numColumns
     */
    public int getNumColumns() {
        return numColumns;
    }

    /**
     * @param numColumns the numColumns to set
     */
    public void setNumColumns(int numColumns) {
        this.numColumns = numColumns;
    }

    /**
     * @return the nameSheetTab
     */
    public String getNameSheetTab() {
        return nameSheetTab;
    }

    /**
     * @param nameSheetTab the nameSheetTab to set
     */
    public void setNameSheetTab(String nameSheetTab) {
        this.nameSheetTab = nameSheetTab;
    }

    /**
     * @return the xlsxWorkbook
     */
    public HSSFWorkbook getXLSWorkbook() {
        return xlsWorkbook;
    }

    /**
     * @param xlsWorkbook the xlsWorkbook to set
     */
    public void setXLSWorkbook(HSSFWorkbook xlsWorkbook) {
        this.xlsWorkbook = xlsWorkbook;
    }

    /**
     * @return the sheetXLS
     */
    public HSSFSheet getSheetXLS() {
        return sheetXLS;
    }

    /**
     * @param sheetXLS the sheetXLS to set
     */
    public void setSheetXLS(HSSFSheet sheetXLS) {
        this.sheetXLS = sheetXLS;
    }

    /**
     * @return the xlsxWorkbook
     */
    public XSSFWorkbook getXLSXWorkbook() {
        return xlsxWorkbook;
    }

    /**
     * @param xlsxWorkbook the xlsxWorkbook to set
     */
    public void setXLSXWorkbook(XSSFWorkbook xlsxWorkbook) {
        this.xlsxWorkbook = xlsxWorkbook;
    }

    /**
     * @return the sheetXLSX
     */
    public XSSFSheet getSheetXLSX() {
        return sheetXLSX;
    }

    /**
     * @param sheetXLSX the sheetXLS to set
     */
    public void setSheetXLSX(XSSFSheet sheetXLSX) {
        this.sheetXLSX = sheetXLSX;
    }
        
}//fin
