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
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Vector;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import model.CustomTableModel;
import model.ExcelWordModel;
import model.PDFModel;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class ViewsController {
    
    public static final String pathExcelSalida = ClassLoader.getSystemResource("templates/PlantillaSalida.xlsx").getPath();
    public static String pathExcelSalidaUsuario = "";
    private ExcelWordModel ewm = null;
    private PDFModel pdfM = null;
    private final char indexColumn[] = {'A','B','C','D','E','F','G','H','I','J','K','L',
                           'M','N','O','P','Q','R','S','T','U','V','X','Y','Z'};
    public static boolean isPDF = false;
    
    public ViewsController(String pathSalida){
        pathExcelSalidaUsuario = pathSalida;
    }

    public ViewsController(File selectedFile) {
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
                  this.pdfM = new PDFModel(selectedFile);
              }
    }
    
    public static boolean isExcel(String fileToReadname){
        //String fileToReadname = selectedFile.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String Excel2003 = "xls";
            String Excel2007 = "xlsx";
            String pdf = "pdf";
            String docx = "docx";
            String doc = "doc";
            if (Excel2003.equalsIgnoreCase(extension) || Excel2007.equalsIgnoreCase(extension) 
               || docx.equalsIgnoreCase(extension) || doc.equalsIgnoreCase(extension) ) {
                  //ewm = new ExcelWordModel(selectedFile);
                ViewsController.isPDF = false;
              } else if (pdf.equalsIgnoreCase(extension)) {
                  //this.pdf = new PDFModel(selectedFile);
                  ViewsController.isPDF = true;
              }
            return isPDF;
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
    
    
    public static CustomTableModel deleteColumn(JTable tabla, int[] indices){
        CustomTableModel dtm = (CustomTableModel) tabla.getModel();
        int nRow = dtm.getRowCount(), nCol = dtm.getColumnCount();
        Vector tableData = new Vector();
        Vector d = null;
        Vector names = new Vector();
        
        for (int i = 0 ; i< nCol ; i++)
                names.add(tabla.getColumnName(i));
        
        for (int i = indices.length-1; i >=0 ; i--)
            names.remove(i);
        
        for (int i = 0 ; i < nRow ; i++){
            d = new Vector();
            for (int j = 0 ; j < nCol ; j++)
                d.add(dtm.getValueAt(i,j));
        tableData.add(d);
        }
        
        for(int index = 0; index < tabla.getModel().getRowCount(); index++)
            for (int i = indices.length - 1; i >= 0; i--)
                ((Vector)tableData.get(index)).remove(i);
        
 
        return new CustomTableModel(tableData, names);
    }

    /**
     * @return the isPDF
     */
    public boolean isPDF() {
        return ViewsController.isPDF;
    }
    
    public static Vector getDataFromTable(JTable tabla){
        CustomTableModel dtm = (CustomTableModel) tabla.getModel();
        int nRow = dtm.getRowCount(), nCol = dtm.getColumnCount();
        Vector tableData = new Vector();
        Vector d = null;
        
        for (int i = 0 ; i < nRow ; i++){
            d = new Vector();
            for (int j = 0 ; j < nCol ; j++)
                d.add(dtm.getValueAt(i,j));
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
    
    public static Vector sortVector(int indexToCompare, Vector data){
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
    
    public static boolean isString(String str){
        try {
            Double.parseDouble(str);
            return false;
        } catch (NumberFormatException e) {
            return true;
        }   
    }
    
    
    
    
}
