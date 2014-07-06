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
import javax.swing.table.DefaultTableModel;
import model.CustomModel;
import model.ExcelModel;
import model.PDFModel;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class ViewsController {
    
    private ExcelModel os = null;
    private PDFModel ps = null;
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
            if (Excel2003.equalsIgnoreCase(extension) || Excel2007.equalsIgnoreCase(extension) ) {
                os = new ExcelModel(selectedFile);
              } else if (pdf.equalsIgnoreCase(extension)) {
                ps = new PDFModel(selectedFile);
              }
    }
    
    
    public CustomModel fillTableArrayList(){
        ArrayList triplete = os.createArrayList();
        CustomModel dtm = new CustomModel();
        dtm.setColumnCount((int)triplete.get(1));
        dtm.setRowCount((int)triplete.get(0));
        int n = 0;
        for (int i = 0; i < (((ArrayList)triplete.get(2)).size()) ; i++){
            dtm.setValueAt(((ArrayList)triplete.get(2)).get(i), n, (i)%(int)triplete.get(1));
            System.out.println("celda: "+i +" en fila: "+n+" y columna: "+(i)%(int)triplete.get(1) + " Dato: "+((ArrayList)triplete.get(2)).get(i));     
            if((i)%(int)triplete.get(1) == 0)
                 n++;
        }
        //System.out.println("filas: " + dtm.getRowCount()+" Columnas: " + dtm.getColumnCount());
    return dtm;
    }
    
    public CustomModel fillTableVector(){
        Vector data = os.createDataVector();
        Vector headers = new Vector();
        for (int i = 0; i < os.getNumColumns(); i++)
            headers.add(this.indexColumn[i]);
        CustomModel dtm = new CustomModel();
        dtm.setColumnCount(os.getNumColumns());
        dtm.setRowCount(os.getNumRows());
        dtm.setDataVector(data, headers);
    
    return dtm;
    }
    
    
    
}
