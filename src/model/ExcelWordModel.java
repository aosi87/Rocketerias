/* Demo Roketeria para el software mineria de datos
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package model;

import controller.ViewsController;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Vector;
import javax.swing.JFrame;
import javax.swing.ProgressMonitorInputStream;
import junit.framework.Assert;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
                this.numSheetTabs = this.xlsWorkbook.getNumberOfSheets();
                //readXLSFile(Excel2003FileToRead);
              } else if (excel2007.equalsIgnoreCase(extension)) {
                FileInputStream Excel2007FileToRead = new FileInputStream(file);
                this.xlsxWorkbook = new XSSFWorkbook(Excel2007FileToRead);
                this.isXSLX = true;
                this.numSheetTabs = this.xlsxWorkbook.getNumberOfSheets();
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
    
    public Vector createDataVectorXLSX(int num){
        data = new Vector();
        Vector d = null;
        this.sheetXLSX = this.xlsxWorkbook.getSheetAt(num);
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
         } 
         //printVector(data);
         this.setNumColumns(rows);
         this.setNumRows(sheetXLSX.getPhysicalNumberOfRows());
         return data;
}
    
    public Vector createDataVectorXLS(int num){
        data = new Vector();
        Vector d = null;
        this.sheetXLS = this.xlsWorkbook.getSheetAt(num);
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
         this.setNumColumns(rows);
         this.setNumRows(sheetXLS.getPhysicalNumberOfRows());
         return data;
}
    
    public void saveExcel(Vector tableData) throws IOException{
        //System.out.println(ViewsController.pathExcelSalida);
        Vector h = new Vector();
        Vector f = new Vector();
        Vector salida;
        File fis = new File(ViewsController.pathExcelSalida);
        this.xlsxWorkbook = new XSSFWorkbook(new FileInputStream(fis));
        Vector dataExcel = this.createDataVectorXLSX(0);
        this.fillVectorsSalidaCore(dataExcel, h, f);
        if(!new File("C:\\Users\\Elpapo\\Desktop\\Demo.xlsx").exists())
         salida = this.mergeVectoresInicial(tableData, h, f);
        else salida = this.mergeVectoresExistente(tableData, dataExcel, f);
        salida.trimToSize();
        this.createDataSalida(salida);
        //this.printVector(headers);
        //this.printVector(dataExcel);
        //this.printVector(footer);
        //System.err.println(((Vector)dataExcel.get(0)).size());
        //this.readExcel();
    }
    
    private Vector mergeVectoresExistente(Vector tabla, Vector excel, Vector footer){
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

        
        FileOutputStream out = new FileOutputStream("C:\\Users\\Elpapo\\Desktop\\Demo.xlsx");
        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
        System.out.println("Success");
    }

    
    private void createDataSalida1(Vector salida){
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sample sheet");
        Vector d = null;
        int rownum = 0;
        for (int i = 0 ; i < salida.size() ; i++) {
            Row row = sheet.createRow(rownum++);
            d = (Vector)salida.get(i);
            int cellnum = 0;
            for (int j = 0 ; j < d.size(); j++) {
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(d.get(j).toString());
            }
        }
 
        try {
            FileOutputStream out = 
                new FileOutputStream(new File("C:\\Users\\Elpapo\\Desktop\\Demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("Excel written successfully..");
     
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
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
        int count = 1;
        for(int i = 0; i < v1.size(); i++){
            for(int j = 0; j < flag; j++)
                if(((Vector)v1.get(i)).get(j).toString().equalsIgnoreCase("") || ((Vector)v1.get(i)).get(j).toString().equalsIgnoreCase(" ") || ((Vector)v1.get(i)).get(j).toString().isEmpty() ){
                    count++;
                    if( count == flag){
                       v1.removeElementAt(i);
                       System.out.println("Contador: "+i);
                       count=1;
                       
                     } 
                    
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
    
    public void printVector(Vector a){
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
        }
        
    }}