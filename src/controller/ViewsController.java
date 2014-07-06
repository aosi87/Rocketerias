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
    
    private ExcelWordModel ewm = null;
    private PDFModel pdf = null;
    private char indexColumn[] = {'A','B','C','D','E','F','G','H','I','J','K','L',
                           'M','N','O','P','Q','R','S','T','U','V','X','Y','Z'};
    
    public ViewsController(){}

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
                  this.pdf = new PDFModel(selectedFile);
              }
    }
    
    
    
    
    public CustomTableModel fillTableVector(){
        Vector data = null;
        if(ewm.getIsXSLX())
            data = ewm.createDataVectorXLSX();
        else 
            data = ewm.createDataVectorXLS();
        
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
        DefaultTableModel dtm = (DefaultTableModel) tabla.getModel();
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
    
    
}
