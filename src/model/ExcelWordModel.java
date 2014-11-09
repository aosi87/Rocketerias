/* Demo Roketeria para el software mineria de datos
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package model;

import controller.ViewController;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Vector;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
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
    //private ProgressMonitorInputStream Excel2003FileToReadPM;
    //private ProgressMonitorInputStream Excel2007FileToReadPM;
    
    private HSSFWorkbook xlsWorkbook = null; //Office 2003 (xls)
    //private HSSFSheet sheetXLS = null;
    private XSSFWorkbook xlsxWorkbook = null; //Office 2007 (xlsx)
    //private XSSFSheet sheetXLSX = null;
    
    public ExcelWordModel(){}
    
    public ExcelWordModel(HSSFWorkbook xlsFile){
        this.xlsWorkbook = xlsFile;
    }
    
    public ExcelWordModel(XSSFWorkbook xlsxFile){
        this.xlsxWorkbook = xlsxFile;
    }
    /*
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
                this.xlsWorkbook.unwriteProtectWorkbook();
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
    }*/
    
    public ExcelWordModel(File file) throws IOException{
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
                this.numSheetTabs = this.xlsWorkbook.getNumberOfSheets();
                //readXLSFile(Excel2003FileToRead);
              } else if (excel2007.equalsIgnoreCase(extension)) {
                OPCPackage pkg = null;
                FileInputStream Excel2007FileToRead = null;
                try {
                    Excel2007FileToRead = new FileInputStream(file);
                    pkg = OPCPackage.open(file);
                } catch (InvalidFormatException ex) {
                    //System.err.println("FormatoDesconocido");
                    //Logger.getLogger(ExcelWordModel.class.getName()).log(Level.SEVERE, null, ex);
                }
                //this.xlsxWorkbook = new XSSFWorkbook(pkg);
                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToRead);
                this.isXSLX = true;
                this.numSheetTabs = this.xlsxWorkbook.getNumberOfSheets();
                //readXLSXFile(Excel2007FileToRead);
              }else if (docx.equalsIgnoreCase(extension)){
                  System.out.println("Docx");
              } else if (doc.equalsIgnoreCase(extension)){
                  System.out.println("Doc");
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
    
    public Vector createDataVectorXLSX(int num){
        data = new Vector();
        Vector d = null;
        XSSFSheet sheetXLSX = null;
        sheetXLSX = this.xlsxWorkbook.getSheetAt(num);
        XSSFRow row;
        int cells = 0;
        //System.err.println(sheetXLSX.getPhysicalNumberOfRows());
         for ( int i = 0; i <= sheetXLSX.getLastRowNum(); i ++ ){
          d = new Vector();
          row = sheetXLSX.getRow( i );
          if(row != null){
              
          if(cells < row.getLastCellNum())
           cells = row.getLastCellNum();
          for ( int j = 0; j < cells; j++ ){
              XSSFCell cell = row.getCell(j, Row.CREATE_NULL_AS_BLANK);
              //if ( cell == null)
                //  d.add("");
              if ( cell != null)
              switch(cell.getCellType()){
                    case XSSFCell.CELL_TYPE_BLANK:
                        d.add("");
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        d.add(cell.getNumericCellValue());
                        //System.out.print(xcell.getNumericCellValue() + " ");
                        break;
                        
                    case XSSFCell.CELL_TYPE_STRING:
                        if(!cell.getStringCellValue().trim().isEmpty())
                         d.add(cell.getStringCellValue());
                        else d.add("");
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
         } }
         //printVector(data);
         this.setNumColumns(cells);
         this.setNumRows(sheetXLSX.getPhysicalNumberOfRows());
         return data;
}
    
    public Vector createDataVectorXLS(int num){
        data = new Vector();
        Vector d = null;
        HSSFSheet sheetXLS= null;
        sheetXLS = this.xlsWorkbook.getSheetAt(num);
        HSSFRow row = null;
        int rows = Math.max(sheetXLS.getPhysicalNumberOfRows(), sheetXLS.getLastRowNum());
        //System.err.println(sheetXLS.getPhysicalNumberOfRows());
         for ( int i = 0; i < rows; i ++ ){
          d = new Vector();
          row = sheetXLS.getRow( i );
          //int index = Math.min(row.getLastCellNum(), 1);
          if(row != null)
          //if(rows < row.getPhysicalNumberOfCells())
           //rows = row.getPhysicalNumberOfCells();
          //System.out.println("Cells: "+row.getPhysicalNumberOfCells() +" "+row.getLastCellNum());
          for ( int j = 0; j < row.getLastCellNum(); j++ ){
              HSSFCell cell = row.getCell(j,Row.RETURN_BLANK_AS_NULL);
              if ( cell == null)
                  d.add("");
              if ( cell != null)
              switch(cell.getCellType()){
                    case HSSFCell.CELL_TYPE_BLANK:
                        d.add("");
                        //System.out.print(xcell.getStringCellValue() + " ");
                        break;
                        
                    case HSSFCell.CELL_TYPE_NUMERIC:
                        d.add(cell.getNumericCellValue());
                        //System.out.print(xcell.getNumericCellValue() + " ");
                        break;
                        
                    case HSSFCell.CELL_TYPE_STRING:
                        if(!cell.getStringCellValue().trim().isEmpty())
                         d.add(cell.getStringCellValue());
                        else d.add("");
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
         this.setNumColumns(sheetXLS.getRow(0).getPhysicalNumberOfCells());
         this.setNumRows(sheetXLS.getPhysicalNumberOfRows());
         return data;
}
    
    public void saveExcel(Vector tableData) throws IOException, FileNotFoundException, InvalidFormatException{
        String path = ViewController.pathExcelSalidaDefault+"/Rocketerias.xlsx";
        //Vector h = new Vector();
        //Vector f = new Vector();
        //Vector salida;
        //File fis = new File(ViewsController.pathExcelSalida);
        //this.xlsxWorkbook = new XSSFWorkbook(new FileInputStream(fis));
        this.fillWorkBook(tableData,path);
        //this.createDataSalida();
        //this.copyHeaderFormat();
        //Vector dataExcel = this.createDataVectorXLSX(0);
        //this.fillVectorsSalidaCore(dataExcel, h, f);
        //if(!new File("C:\\Users\\Elpapo\\Desktop\\Demo.xlsx").exists())
        //salida = this.mergeVectoresInicial(tableData, h, f);
        //else salida = this.mergeVectoresExistentes(tableData, dataExcel, f);
        //salida.trimToSize();
        //this.createDataSalida(salida);
        //this.printVector(headers);
        //this.printVector(dataExcel);
        //this.printVector(footer);
        //System.err.println(((Vector)dataExcel.get(0)).size());
        //this.readExcel();
    }
    
    private void fillWorkBook(Vector tableData, String path) throws FileNotFoundException, IOException, InvalidFormatException{
            //this.xlsxWorkbook.setSheetName(0, "Proshop");
           SXSSFWorkbook wb = new SXSSFWorkbook(500);
            
           this.createHeader(wb);
           if(this.checkFile(path)){
                File f = new File(path); 
                OPCPackage pkg = pkg = OPCPackage.open(f);
                this.xlsxWorkbook = new XSSFWorkbook(pkg);
                Vector oldData = this.createDataVectorXLSX(0);
                for(int i = 5; i >= 0; i--)
                    oldData.removeElementAt(i);
                for(int i = 0; i < 7; i++)
                    oldData.removeElementAt(oldData.size()-1);
                this.fillWithTableData(wb, oldData);
                pkg.close();
                //System.out.println(path);
           
           }
           this.fillWithTableData(wb, tableData);
           this.createFooter(wb);
           this.setColumnWidth(wb.getSheetAt(0));
           this.setCombinedCells(wb.getSheetAt(0));
           //sheet.shiftRows(5, 6, 10);
                                            //col, fil
           wb.getSheetAt(0).createFreezePane(0, 6);
           this.addImage(wb);
           this.createDataSalida(wb, path);
           
    }
    
    private void fillWithTableData(SXSSFWorkbook wb, Vector tableData){
        Sheet sheet = wb.getSheetAt(0);
        Row row = null;
        Cell c = null;
        int index = sheet.getLastRowNum()+1;
        for(int i = 0 ; i < tableData.size(); i++){
            row = sheet.createRow(index+i);
            for(int j = 0; j < ((Vector)tableData.get(i)).size() ; j++ ){
                c = row.createCell(j);
                if(this.isString(((Vector)tableData.get(i)).get(j).toString()))
                    c.setCellValue(((Vector)tableData.get(i)).get(j).toString());
                else c.setCellValue(Double.parseDouble(((Vector)tableData.get(i)).get(j).toString()));
            }    
        }
    }
    
    private boolean isString(String str){
        try {
            Double.parseDouble(str);
            return false;
        } catch (NumberFormatException e) {
            return true;
        }   
    }
    
    private void setColor(CellStyle cs, IndexedColors color){
        cs.setFillForegroundColor(color.index);
        cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
    }
    
    private Vector mergeVectoresExistentes(Vector tabla, Vector excel, Vector footer){
        Vector salida = new Vector();
        for(int i = 0; i < excel.size()-7; i++)
            salida.add(excel.get(i));
        for(int i = 0; i < tabla.size(); i++)
            salida.add(tabla.get(i));
        for(int i = 0; i < footer.size(); i++)
            salida.add(footer.get(i));
        this.printVector(salida);
        return salida;
    }
    
    private void createDataSalida(Vector salida) throws FileNotFoundException, IOException{
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        Sheet sh = wb.createSheet();
        Vector d = new Vector(); 
        for(int rownum = 0; rownum < salida.size(); rownum++){
            Row row = sh.createRow(rownum);
            d = (Vector)salida.get(rownum);
            for(int cellnum = 0; cellnum < d.size() ; cellnum++){
                Cell cell = row.createCell(cellnum);
                //String address = new CellReference(cell.toString()).formatAsString();
                if(d.get(cellnum) != null)
                cell.setCellValue(d.get(cellnum).toString());
                else cell.setCellValue(Cell.CELL_TYPE_BLANK);
            }
            // manually control how rows are flushed to disk 
           if(rownum % 100 == 0) {
                ((SXSSFSheet)sh).flushRows(100); // retain 100 last rows and flush all others

                // ((SXSSFSheet)sh).flushRows() is a shortcut for ((SXSSFSheet)sh).flushRows(0),
                // this method flushes all rows
           }

        }

        
        FileOutputStream out = new FileOutputStream("");
        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
        //System.out.println("Success");
    }
    
    public void createDataSalida(SXSSFWorkbook xlsxWB, String path) throws FileNotFoundException, IOException{
//        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
//        Sheet sh = wb.createSheet();
//        wb.setSheetName(0, xlsxWB.getSheetName(0));
//        XSSFSheet oldsh  = xlsxWB.getSheetAt(0);
//        Row d , row = null; 
//        Cell cell = null;
//        //for(int rownum = 0; rownum <  6; rownum++){
//        for(int rownum = 0; rownum <  oldsh.getPhysicalNumberOfRows(); rownum++){
//            row = sh.createRow(rownum);
//            d = oldsh.getRow(rownum);
//            for(int cellnum = 0; cellnum < d.getPhysicalNumberOfCells() ; cellnum++){
//                 cell = row.createCell(cellnum);
//                //String address = new CellReference(cell.toString()).formatAsString();
//                if(d.getCell(cellnum) != null){
//                cell.setCellValue(d.getCell(cellnum).toString());
//                cell.setCellStyle(d.getCell(cellnum).getCellStyle());
//                } else cell.setCellValue(Cell.CELL_TYPE_BLANK);
//            }
//            // manually control how rows are flushed to disk 
//           if(rownum % 100 == 0) {
//                ((SXSSFSheet)sh).flushRows(100); // retain 100 last rows and flush all others
//
//                // ((SXSSFSheet)sh).flushRows() is a shortcut for ((SXSSFSheet)sh).flushRows(0),
//                // this method flushes all rows
//           }
//
//        }

        
        FileOutputStream out = new FileOutputStream(path);//"C:\\Users\\Elpapo\\Desktop\\Demo.xlsx");
        xlsxWB.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        xlsxWB.dispose();
        //System.out.println("Success");
    }
    
    private Vector mergeVectoresInicial(Vector tabla, Vector headers, Vector footer){
        Vector salida = new Vector();
        for(int i = 0; i < headers.size(); i++)
            salida.add(headers.get(i));
        for(int i = 0; i < tabla.size(); i++)
            salida.add(tabla.get(i));
        for(int i = 0; i < footer.size(); i++)
            salida.add(footer.get(i));
        this.printVector(salida);
        return salida;
    }
    
    private void fillVectorsSalidaCore(Vector salida, Vector v1, Vector v2) {
        for(int i = 0; i < 6; i++)//salida.size()-7; i++)
            v1.add(salida.get(i));
        for(int j = salida.size()-7; j < salida.size(); j++)
            v2.add(salida.get(j));
        //this.cleanVector(salida);
    }
    
    private void cleanVector(Vector v1){
        int flag = ((Vector)v1.get(0)).size();
        for(int i = 0; i < v1.size(); i++){
            for(int j = 0; j < flag; j++)
                if(((Vector)v1.get(i)).get(j).toString().equalsIgnoreCase("") || ((Vector)v1.get(i)).get(j).toString().equalsIgnoreCase(" ") || ((Vector)v1.get(i)).get(j).toString().isEmpty() ){
                    v1.removeElementAt(i);                       
                 } 
                    
            }                   
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
    public String getNameSheetTab(int i) {
        if(this.isXSLX){
            this.nameSheetTab = this.xlsxWorkbook.getSheetName(i);
        }
        else {
            this.nameSheetTab = this.xlsWorkbook.getSheetName(i);
        }
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
     *
    public HSSFSheet getSheetXLS() {
        return sheetXLS;
    }*/

    /**
     * @param sheetXLS the sheetXLS to set
     *
    public void setSheetXLS(HSSFSheet sheetXLS) {
        this.sheetXLS = sheetXLS;
    }*/

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
     *
    public XSSFSheet getSheetXLSX() {
        return sheetXLSX;
    }*/

    /**
     * @param sheetXLSX the sheetXLS to set
     *
    public void setSheetXLSX(XSSFSheet sheetXLSX) {
        this.sheetXLSX = sheetXLSX;
    }*/
    
    public void printArrayList(ArrayList a){
                             //filas   columnas
        System.out.println(a.get(0)+" "+a.get(1));
        for (int i = 0 ; i < ((ArrayList)a.get(2)).size(); i++){
          //if(!((ArrayList)a.get(2)).get(i).toString().equalsIgnoreCase(""))
              System.out.print(((ArrayList)a.get(2)).get(i).toString());
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
    
    public static void printVector(Vector a){
        try{                     //filas   columnas
        System.out.println(a.size()+" "+((Vector)a.get(1)).size());
        } catch (ArrayIndexOutOfBoundsException e){e.printStackTrace();}
        for (int i = 0 ; i < a.size(); i++){
          //if(!((ArrayList)a.get(2)).get(i).toString().equalsIgnoreCase(""))
              System.out.println(a.get(i));
          //if((i+1)%(int)a.get(1) == 0)
             //System.out.println();
        }   
    } 

   public static void main(String[] args) {
        
            //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\plantilla salida.xlsx");
            //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Plantillas_salida_Listasprecios.xlsx");
            //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\EAW 2013 Distributor Pricelist AUGUST Update.xlsx");
            //File fileToRead = new File("\\Users\\Manuu Alcocer\\Downloads\\Statistics 6 T-test.xlsx");
            //File fileToRead = new File("C:\\Users\\Elpapo\\Documents\\NetBeansProjects\\Rocketerias\\resources\\templates\\plantilla salida.xlsx");
            
            ExcelWordModel ou = new ExcelWordModel();
        try {
        //    ou = new ExcelWordModel(new File(ViewsController.pathExcelSalida));
            ou.saveExcel(null);
        } catch (IOException ex) {
            //Logger.getLogger(ExcelWordModel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            //Logger.getLogger(ExcelWordModel.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
   
    private void setCombinedCells(Sheet sheet){
            int rows = sheet.getPhysicalNumberOfRows();
                                //rowFrom,rowTo,colFrom,colTo
            sheet.addMergedRegion(new CellRangeAddress(3,3,2,6));
            sheet.addMergedRegion(new CellRangeAddress(4,4,2,6));
            sheet.addMergedRegion(new CellRangeAddress(rows-1,rows-1,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-2,rows-2,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-3,rows-3,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-4,rows-4,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-5,rows-5,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-6,rows-6,0,11));
            sheet.addMergedRegion(new CellRangeAddress(rows-7,rows-7,0,11));            
    }

    private void setColumnWidth(Sheet sheet) {
        final int colWidth = 544;
        sheet.setColumnWidth(0, colWidth*16);
        sheet.setColumnWidth(1, colWidth*13);
        for(int i = 2; i < 14; i++)
        sheet.setColumnWidth(i, colWidth*5);
    }
    
    private void createHeader(SXSSFWorkbook wb){
            String cabecera[] = { 
               "Modelo",
                "Descripción",
                "No.Parte",
                "Codigo de Barras",
                "Unidad de Medida",
                "Costo",
                "Precio Publico",
                "Precio Contratista",
                "Precio Mayoreo",
                "Precio Super Mayoreo",
                "Stock",
                "Imagen",
                "Giro",
                "Marca"      };
            
            wb.createSheet("ProShop");
            Font font = wb.createFont();
            Sheet sheet = wb.getSheetAt(0);//this.xlsxWorkbook.getSheetAt(0);
            Row row = null;
            Cell c = null;
            XSSFCellStyle cs = (XSSFCellStyle)wb.createCellStyle();
            this.setColor(cs, IndexedColors.GREY_25_PERCENT);
            
            for(int i = 0; i < 5; i++){
             row = sheet.createRow(i);
             for (int j = 0 ; j < 14; j++){
              c = row.createCell(j);
              c.setCellStyle(cs);
              c.setCellType(Cell.CELL_TYPE_STRING);
              c.setCellValue("");
             }
            }
            //Fill cells manually -first text
            row = sheet.getRow(3);
            c = row.getCell(2);
            c.setCellValue("Lista De precios de Rocketerias Distribuidores S.A de C.V.");
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            font.setColor(IndexedColors.RED.index);
            //next text
            row = sheet.getRow(4);
            c = row.getCell(2);
            c.setCellValue("Confidencial.");
            cs = (XSSFCellStyle)wb.createCellStyle();
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            this.setColor(cs, IndexedColors.GREY_25_PERCENT);
            cs.setFont(font);
            c.setCellStyle(cs);
            //Create headers
            row = sheet.createRow(5);
            cs = (XSSFCellStyle)wb.createCellStyle();
            this.setColor(cs, IndexedColors.BLACK);
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            cs.setWrapText(true);
            font = wb.createFont();
            font.setColor(IndexedColors.WHITE.getIndex());
            font.setFontHeightInPoints((short)9);
            font.setFontName("Arial Bold");
            cs.setFont(font);
            for(int i = 0; i < cabecera.length ; i++){
                c = row.createCell(i);
                 c.setCellType(Cell.CELL_TYPE_STRING);
                 c.setCellValue(cabecera[i]);
                 c.setCellStyle(cs);
            } 
    }
    
    private void createFooter(SXSSFWorkbook wb){
            
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            XSSFCellStyle cs = null;
            Font font = null;
            Cell c = null;
            int rows = sheet.getPhysicalNumberOfRows();
            
        //create footer
            row = sheet.createRow(rows++);
            //c = row.createCell(0);
            cs = (XSSFCellStyle)wb.createCellStyle();
            this.setColor(cs, IndexedColors.RED);
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            cs.setWrapText(true);
            font = wb.createFont();
            font.setColor(IndexedColors.WHITE.getIndex());
            font.setFontHeightInPoints((short)11);
            font.setFontName("Arial Bold");
            cs.setFont(font);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            } 
            c = row.getCell(0);
            c.setCellValue("PRECIOS EN (Euros, Dolares, pesos, Libra esterlina) + IVA");
            c.setCellType(Cell.CELL_TYPE_STRING);
            
            //next data
            row = sheet.createRow(rows++);
            cs = (XSSFCellStyle)wb.createCellStyle();
            //this.setColor(cs, IndexedColors.RED);
            cs.setAlignment(HorizontalAlignment.CENTER);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            cs.setWrapText(true);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue("ROCKETERIAS DISTRIBUIDORES, S. A. de C. V.");
            c.setCellType(Cell.CELL_TYPE_STRING);
            
            //next data
            row = sheet.createRow(rows++);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue("Oficina Matriz: Calle 21 No. 802 X 64 Col. Jardines de Mérida");
            c.setCellType(Cell.CELL_TYPE_STRING);
            
            //next data
            row = sheet.createRow(rows++);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue(" Tel: (52) (999) 930-06-00 ");
            c.setCellType(Cell.CELL_TYPE_STRING);
            
            //next data
            row = sheet.createRow(rows++);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue("Mérida, Yucatán, México  C. P. 97135");
            c.setCellType(Cell.CELL_TYPE_STRING);
            
            //next data
            row = sheet.createRow(rows++);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue("info@rocketerias.com.mx");
            c.setCellType(Cell.CELL_TYPE_STRING);
                        
            //next data
            row = sheet.createRow(rows++);
            for(int i = 0; i < 14 ;i++ ){
             c = row.createCell(i);
             c.setCellStyle(cs);
            }
            c = row.getCell(0);
            c.setCellValue("*Precios sujetos a cambios sin previo aviso.");
            c.setCellType(Cell.CELL_TYPE_STRING);
        
    }
    
    private void setThinBorder(CellStyle style){
        style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        style.setBorderTop(XSSFCellStyle.BORDER_THIN);
        style.setBorderRight(XSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
    }
    
    private void addImage(SXSSFWorkbook wb) throws IOException{
         Sheet sheet = wb.getSheetAt(0);
         //FileInputStream obtains input bytes from the image file
         FileInputStream inputStream = new FileInputStream(getClass().getClassLoader().getResource("gfx/logoExcel.png").getPath());
         //Get the contents of an InputStream as a byte[].
         byte[] bytes = IOUtils.toByteArray(inputStream);
         //Adds a picture to the workbook
         int pictureIdx = wb.addPicture(bytes, wb.PICTURE_TYPE_PNG);
         //close the input stream
         inputStream.close();
 
        //Returns an object that handles instantiating concrete classes
         CreationHelper helper = wb.getCreationHelper();
 
        //Creates the top-level drawing patriarch.
         Drawing drawing = sheet.createDrawingPatriarch();
 
        //Create an anchor that is attached to the worksheet
         ClientAnchor anchor = helper.createClientAnchor();
         //set top-left corner for the image
         anchor.setCol1(0);
         anchor.setRow1(1);
 
         //Creates a picture
         Picture pict = drawing.createPicture(anchor, pictureIdx);
         //Reset the image to the original size
         pict.resize();
    }

    private boolean checkFile(String path) {
        if(new File(path).exists())
            return true;
        else return false;
    }
}