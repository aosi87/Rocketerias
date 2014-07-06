/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package model;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Vector;
import javax.swing.JFrame;

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
public class ExcelWordModel {
    
    private boolean isXSLX = false;
    private int numSheetTabs = 0;
    private int numRows = 0;
    private int numColumns = 0;
    private String nameSheetTab = "";
    private Vector data = null;
    private ProgressMonitorInputStream Excel2003FileToReadPM;
    private ProgressMonitorInputStream Excel2007FileToReadPM;
    
    private HSSFWorkbook xlsWorkbook = null; //Office 2003 (xls)
    private HSSFSheet sheetXLS = null;
    private XSSFWorkbook xlsxWorkbook = null; //Office 2007 (xlsx)
    private XSSFSheet sheetXLSX = null;
    
    public ExcelWordModel(){}
    
    public ExcelWordModel(HSSFWorkbook xlsFile){
        this.xlsWorkbook = xlsFile;
    }
    
    public ExcelWordModel(XSSFWorkbook xlsxFile){
        this.xlsxWorkbook = xlsxFile;
    }
    
    public ExcelWordModel(File file, JFrame frame){
            try {
            String fileToReadname = file.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String excel2003 = "xls";
            String excel2007 = "xlsx";
            if (excel2003.equalsIgnoreCase(extension)) {
                Excel2003FileToReadPM = new ProgressMonitorInputStream(frame,"Leyendo",new FileInputStream(file));
                //Excel2003FileToReadPM = new FileInputStream(file);
                this.xlsWorkbook = new HSSFWorkbook(Excel2003FileToReadPM);
                this.isXSLX = false;
                readXLSFile();
              } else if (excel2007.equalsIgnoreCase(extension)) {
                Excel2007FileToReadPM = new ProgressMonitorInputStream(frame,"Leyendo",new FileInputStream(file));  
                //Excel2007FileToRead = new FileInputStream(file);
                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToReadPM);
                this.isXSLX = true;
                readXLSXFile();
              }
            } catch (IOException ex) {
                System.out.println(ex.getMessage());
              }
    }
    
    public ExcelWordModel(File file){
            try {
            String fileToReadname = file.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String excel2003 = "xls";
            String excel2007 = "xlsx";
            String docx = "docx";
            String doc = "doc";
            if (excel2003.equalsIgnoreCase(extension)) {
                FileInputStream Excel2003FileToRead = new FileInputStream(file);
                this.xlsWorkbook = new HSSFWorkbook(Excel2003FileToRead);
                this.isXSLX = false;
                //readXLSFile(Excel2003FileToRead);
              } else if (excel2007.equalsIgnoreCase(extension)) {
                FileInputStream Excel2007FileToRead = new FileInputStream(file);
                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToRead);
                this.isXSLX = true;
                //readXLSXFile(Excel2007FileToRead);
              }else if (docx.equalsIgnoreCase(extension)){
                  System.out.println("Docx");
              } else if (doc.equalsIgnoreCase(extension)){
                  System.out.println("Doc");
              }
            } catch (IOException ex) {
                System.out.println(ex.getMessage());
              }
    }
    
    /**
     * Con este metodo leeremos los archivos en modo de compatibilidad, es decir los archivos
     * con extencion .xls o creados con Microsoft Office 2007 o anteriores.
     * 
     * @throws java.io.IOException 
     */
    public void readXLSFile() throws IOException {
        HSSFWorkbook wb = this.xlsWorkbook; //new HSSFWorkbook(Excel2003FileToRead);
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
     * @throws IOException 
     */
    public void readXLSXFile() throws IOException{
        XSSFWorkbook xwb = this.xlsxWorkbook;//new XSSFWorkbook(Excel2007FileToRead);
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
    
    public Vector createDataVectorXLSX(){
        data = new Vector();
        Vector d = null;
        this.sheetXLSX = this.xlsxWorkbook.getSheetAt(0);
        XSSFRow row;
        int rows = 0;
        System.err.println(sheetXLSX.getPhysicalNumberOfRows());
         for ( int i = 0; i < sheetXLSX.getPhysicalNumberOfRows(); i ++ ){
          d = new Vector();
          row = sheetXLSX.getRow( i );
          if(rows < row.getPhysicalNumberOfCells())
           rows = row.getPhysicalNumberOfCells();
          for ( int j = 0; j < row.getPhysicalNumberOfCells(); j++ ){
              XSSFCell cell = row.getCell( j );
              if ( cell == null)
                  d.add("");
              if ( cell != null)
              switch(cell.getCellType()){
                    case XSSFCell.CELL_TYPE_BLANK:
                        d.add(" ");
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        d.add(cell.getNumericCellValue());
                        //System.out.print(xcell.getNumericCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_STRING:
                        d.add(cell.getStringCellValue());
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    default:
                        d.add("NULL");
                        //if(xcell.getRowIndex() == numColumn)  
                        //System.out.print("");
                        break;
                }
          }
                    //d.add( "\n" );
                    data.add( d );
         } 
         //printVector(data);
         this.setNumColumns(rows);
         this.setNumRows(sheetXLSX.getPhysicalNumberOfRows());
         return data;
}
    
    public Vector createDataVectorXLS(){
        data = new Vector();
        Vector d = null;
        this.sheetXLS = this.xlsWorkbook.getSheetAt(0);
        HSSFRow row;
        int rows = 0;
        System.err.println(sheetXLS.getPhysicalNumberOfRows());
         for ( int i = 0; i < sheetXLS.getPhysicalNumberOfRows(); i ++ ){
          d = new Vector();
          row = sheetXLS.getRow( i );
          if(rows < row.getPhysicalNumberOfCells())
           rows = row.getPhysicalNumberOfCells();
          for ( int j = 0; j < row.getPhysicalNumberOfCells(); j++ ){
              HSSFCell cell = row.getCell( j );
              if ( cell == null)
                  d.add("");
              if ( cell != null)
              switch(cell.getCellType()){
                    case HSSFCell.CELL_TYPE_BLANK:
                        d.add(" ");
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    case HSSFCell.CELL_TYPE_NUMERIC:
                        d.add(cell.getNumericCellValue());
                        //System.out.print(xcell.getNumericCellValue() + " ");
                        break;
                        
                    case HSSFCell.CELL_TYPE_STRING:
                        d.add(cell.getStringCellValue());
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    default:
                        d.add("NULL");
                        //if(xcell.getRowIndex() == numColumn)  
                        //System.out.print("");
                        break;
                }
          }
                    //d.add( "\n" );
                    data.add( d );
         } 
         //printVector(data);
         this.setNumColumns(rows);
         this.setNumRows(sheetXLS.getPhysicalNumberOfRows());
         return data;
}

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
    
    public void printArrayList(ArrayList a){
                             //filas   columnas
        System.out.println(a.get(0)+" "+a.get(1));
        for (int i = 0 ; i < ((ArrayList)a.get(2)).size(); i++){
          //if(!((ArrayList)a.get(2)).get(i).toString().equalsIgnoreCase(""))
              System.out.print(((ArrayList)a.get(2)).get(i).toString()+" ");
          if((i+1)%(int)a.get(1) == 0)
              System.out.println(i+"-"+(i+1)%(int)a.get(1));
        }   
    }
    
        /**
     * @return the isXSLX
     */
    public boolean getIsXSLX() {
        return isXSLX;
    }

    /**
     * @param isXSLX the isXSLX to set
     */
    public void setIsXSLX(boolean isXSLX) {
        this.isXSLX = isXSLX;
    }
    
    public void printVector(Vector a){
                             //filas   columnas
        System.out.println(a.size()+" "+((Vector)a.get(1)).size());
        for (int i = 0 ; i < a.size(); i++){
          //if(!((ArrayList)a.get(2)).get(i).toString().equalsIgnoreCase(""))
              System.out.println(a.get(i));
          //if((i+1)%(int)a.get(1) == 0)
           //   System.out.println();
        }   
    } 

  /* public static void main(String[] args) {
        //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\plantilla salida.xlsx");
        //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Plantillas_salida_Listasprecios.xlsx");
        //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\EAW 2013 Distributor Pricelist AUGUST Update.xlsx");
        File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Statistics 6 T-test.xlsx");
        ExcelWordModel ou;
//        JFrame jf = new JFrame();
//        jf.setSize(400, 400);
//        jf.setVisible(true);
//        jf.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        ou = new ExcelWordModel(fileToRead, jf);
        //ou.createArrayList();
        ou = new ExcelWordModel(fileToRead);
        ou.createDataVectorXLSX();
    }*/
}