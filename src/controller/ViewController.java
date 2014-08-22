/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Scanner;
import java.util.Vector;
import javax.swing.JTable;
import model.CustomTableModel;
import model.ExcelWordModel;
import model.PDFModel;
import org.nerdpower.tabula.RectangularTextContainer;
import org.nerdpower.tabula.Table;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class ViewController {
    
    public static final String pathExcelTemplate = "resources/templates/plantillasalida.xlsx";
    public static String outPutFileName = "extractosTablas.csv";
    public static String pathExcelSalidaDefault = System.getProperty("user.home") + System.getProperty("file.separator")+"Desktop";
    public static String pathExcelSalidaUsuario ="";
    public static boolean isCSV = false;
    public static boolean isExcel;
    private ExcelWordModel ewm = null;
    private PDFModel pdfM = null;
    private final char indexColumn[] = {'A','B','C','D','E','F','G','H','I','J','K','L',
                           'M','N','O','P','Q','R','S','T','U','V','X','Y','Z'};
    public static boolean isPDF = false;
    
    public ViewController(String pathSalida){
        pathExcelSalidaUsuario = pathSalida;
    }
    
    public ViewController(){}

    public ViewController(File selectedFile) throws IOException {
            String fileToReadname = selectedFile.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String Excel2003 = "xls";
            String Excel2007 = "xlsx";
            String pdf = "pdf";
            String docx = "docx";
            String doc = "doc";
            if (Excel2003.equalsIgnoreCase(extension) || Excel2007.equalsIgnoreCase(extension) 
               || docx.equalsIgnoreCase(extension) || doc.equalsIgnoreCase(extension) ) {
                ewm = new ExcelWordModel(selectedFile);
              } else if (pdf.equalsIgnoreCase(extension)) {
                  //this.pdfM = new PDFModel(selectedFile);
              }
    }
    
    public static void setPDFCSVExcel(String fileToReadname){
        //String fileToReadname = selectedFile.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String Excel2003 = "xls";
            String Excel2007 = "xlsx";
            String pdf = "pdf";
            String docx = "docx";
            String doc = "doc";
            String csv = "csv";
            if (Excel2003.equalsIgnoreCase(extension) || Excel2007.equalsIgnoreCase(extension) 
               || docx.equalsIgnoreCase(extension) || doc.equalsIgnoreCase(extension)) {
                  //ewm = new ExcelWordModel(selectedFile);
                ViewController.isPDF = false;
                ViewController.isExcel = true;
                ViewController.isCSV = false;
              } else if (pdf.equalsIgnoreCase(extension)) {
                  //this.pdf = new PDFModel(selectedFile);
                  ViewController.isCSV = false;
                  ViewController.isExcel = false;
                  ViewController.isPDF = true;
              } else {
                  ViewController.isCSV = true;
                  ViewController.isExcel = false;
                  ViewController.isPDF = false;
              }
    }
    
    public int getNumSheet(){
        return ewm.getNumSheetTabs();
    }
    
    public String getNameSheet(int i){
        return ewm.getNameSheetTab(i);
    }   
    
    
    
    public CustomTableModel fillTableVector(int sheetNumber){
        Vector data = null;
        if(ewm.getIsXSLX())
            data = ewm.createDataVectorXLSX(sheetNumber);
        else 
            data = ewm.createDataVectorXLS(sheetNumber);
        
        Vector headers = new Vector();
        
        for (int i = 0; i < ewm.getNumColumns(); i++)
            headers.add(this.indexColumn[i]);
        
        CustomTableModel dtm = new CustomTableModel();
        dtm.setColumnCount(ewm.getNumColumns());
        dtm.setRowCount(ewm.getNumRows());
        dtm.setDataVector(data, headers);
    
    return dtm;
    }
    
    
    public static /*CustomTableModel*/void deleteColumn(JTable tabla, int[] indices){
        CustomTableModel dtm = (CustomTableModel) tabla.getModel();
        int nRow = dtm.getRowCount(), nCol = dtm.getColumnCount();
        Vector tableData = new Vector();
        Vector d = null;
        Vector names = new Vector();
        
        for (int i = 0 ; i < nCol ; i++)
                names.add(tabla.getColumnName(i));
        
        for (int i = indices.length-1; i >= 0 ; i--)
            names.remove(indices[i]);
        
//        System.out.println("indices: "+indices.length);
//        System.out.println("Names: "+names);
        
        for (int i = 0 ; i < nRow ; i++){
            d = new Vector();
            for (int j = 0 ; j < nCol ; j++)
                d.add(dtm.getValueAt(i,j));
        tableData.add(d);
        }
        
        for(int index = 0; index < tabla.getModel().getRowCount(); index++)
            for (int i = indices.length - 1; i >= 0; i--)
                ((Vector)tableData.get(index)).remove(indices[i]);
        

        dtm.setDataVector(tableData, names);
        //return dtm;
    }

    /**
     * @return the isPDF
     */
    public boolean isPDF() {
        return ViewController.isPDF;
    }
    
    public static Vector getDataFromTable(JTable tabla){
        CustomTableModel dtm = (CustomTableModel) tabla.getModel();
        int nRow = dtm.getRowCount(), nCol = dtm.getColumnCount();
        Vector tableData = new Vector();
        Vector d = null;
        
        for (int i = 0 ; i < nRow ; i++){
            d = new Vector();
            for (int j = 0 ; j < nCol ; j++)
                if(dtm.getValueAt(i,j) == null)
                    d.add("");
                else  d.add(dtm.getValueAt(i,j));
        tableData.add(d);
        }
        return tableData;
    }
    
    public static void saveData(int indexToCompare, JTable tabla) throws Exception{
        Vector data = getDataFromTable(tabla);
        //System.out.println("Salvando datos."+ data);
        new ExcelWordModel().saveExcel(data);
        //sortVector(indexToCompare,data);
    }
    
    public static Vector sortVector(final int indexToCompare, Vector data){
        //List newData = new Vector();
        //data.sort(Comparator.naturalOrder());
        Collections.sort(data, new Comparator<Vector<String>>(){
          @Override  
          public int compare(Vector<String> v1, Vector<String> v2) {
              return v1.get(indexToCompare).compareTo(v2.get(indexToCompare)); //If you order by 2nd element in row
            }});
        //System.out.print("Ordenando.."+data);
        return null;//(Vector)newData;
    }
    
    public static void createCSV( String path ){
        new Tabula(new String[]{path,"-n", "-p all", "-o " + ViewController.outPutFileName});
    }
    
    public static CustomTableModel fillVectorCSV( String path ){
        
        Tabula tb = new Tabula(new String[]{path,"-n", "-p all", "-o " + ViewController.outPutFileName});
        List<Table> tables = tb.getTable();
        Vector<Vector> data = new Vector();
        
        for(Table table : tables)
         for (List<RectangularTextContainer> row: table.getRows()) {
            Vector<String> rows = new Vector();
            for (RectangularTextContainer tc: row) {
                rows.add(tc.getText());    
            }
            data.add(rows);
         }
            
        return new CustomTableModel(data, null);
    }
    
    public static boolean isString(String str){
        try {
            Double.parseDouble(str);
            return false;
        } catch (NumberFormatException e) {
            return true;
        }   
    }
    
    public static CustomTableModel populateTableCSV(String source) throws FileNotFoundException{
      CustomTableModel model = new CustomTableModel();
      int col = 0;
      int fila = 0;
      Scanner scan = new Scanner(new File(source));
      String[] array;
      Vector<Vector> data = new Vector();
      Vector<String> d = null;
      while (scan.hasNextLine()) {
        d = new Vector();
        String line = scan.nextLine();
        if(line.indexOf(",")>-1)
            array = line.split(",");
        else
            array = line.split("\t");
        
        for (int i = 0; i < array.length ; i ++){
            d.add(array[i]);
        }
        data.add(d);
        }
      for (int i = 0 ; i < data.size() ; i++){
          for(int j = 0 ; j < ((Vector)data.get(i)).size();  j++){
                    if(col < ((Vector)data.get(i)).size() )
                        col = ((Vector)data.get(i)).size();
            }
      }
      
      model.setRowCount(data.size());
      model.setColumnCount(col);//((Vector)data.get(0)).size());
      //model.setDataVector(data, new Vector());
      System.out.println(model.getRowCount()+" -- "+ model.getColumnCount() +": col = "+ col);
      for (int i = 0 ; i < model.getRowCount() ; i++){
          for(int j = 0 ; j < model.getColumnCount();  j++){
             if(col > ((Vector)data.get(i)).size())
              ((Vector)data.get(i)).add("");
             else
               //System.out.println(i+".- "+model.getValueAt(i, j));
               model.setValueAt(((Vector)data.get(i)).get(j), i, j);
              }
      }
      
      return model;
} 
    
    public static CustomTableModel insertDataCSV(String source) throws FileNotFoundException {
        CustomTableModel model = new CustomTableModel();
        Scanner scan = new Scanner(new File(source));
        String[] array;
        while (scan.hasNextLine()) {
            String line = scan.nextLine();
            if(line.contains(","))
                array = line.split(",");
            else
                array = line.split("\t");
          Object[] data = new Object[array.length];
            System.arraycopy(array, 0, data, 0, array.length);
        model.setColumnIdentifiers(new Vector());
        model.addRow(data);
    }
    return model;
} 
    
    
    
}
