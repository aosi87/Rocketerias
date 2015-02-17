/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

import au.com.bytecode.opencsv.CSVReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Scanner;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JComboBox;
import javax.swing.JTable;
import model.CustomTableModel;
import model.ExcelWordModel;
import model.PDFModel;
import org.apache.commons.cli.ParseException;
import org.nerdpower.tabula.RectangularTextContainer;
import org.nerdpower.tabula.Table;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class ViewController {
    
    //public final String pathExcelTemplate = getClass().getClassLoader().getResource("templates/plantillasalida.xlsx").getPath();
    public static String pathInicioUsuario = "extractoPDF.csv";
    public static String pathExcelSalidaDefault = System.getProperty("user.home") + System.getProperty("file.separator")+"Desktop";
    public static String pathExcelSalidaUsuario ="";
    public static boolean isCSV = false;
    public static boolean isExcel;
    public static String[] marcas;
    public static float[] factores;

    private ExcelWordModel ewm = null;
    private PDFModel pdfM = null;
    private static final char indexColumn[] = {'A','B','C','D','E','F','G','H','I','J','K','L',
                           'M','N','O','P','Q','R','S','T','U','V','X','Y','Z'};
    public static boolean isPDF = false;
    
    public ViewController(String pathSalida){
        pathExcelSalidaUsuario = pathSalida;
    }
    
    public ViewController(){}

    public ViewController(File selectedFile) throws IOException {
            ViewController.cargarMarcas();
            String fileToReadname = selectedFile.getName();
            String extension = fileToReadname.substring(fileToReadname.lastIndexOf(".")
                    + 1, fileToReadname.length());
            String Excel2003 = "xls";
            String Excel2007 = "xlsx";
            String pdf = "pdf";
            String docx = "docx";
            String doc = "doc";
            //System.out.println(Arrays.toString(ViewController.marcas));
            if (Excel2003.equalsIgnoreCase(extension) || Excel2007.equalsIgnoreCase(extension) 
               || docx.equalsIgnoreCase(extension) || doc.equalsIgnoreCase(extension) ) {
                ewm = new ExcelWordModel(selectedFile);
              } else if (pdf.equalsIgnoreCase(extension)) {
                  //this.pdfM = new PDFModel(selectedFile);
              }
    }
    
    public static void cargarMarcas() {
        try {
            ViewController.marcas = readCVSMarcas();
        } catch (FileNotFoundException ex) {
            System.out.println("No se pudo leer marcas.txt");
            //Logger.getLogger(ViewController.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private static String[] readCVSMarcas() throws FileNotFoundException {
      Scanner scan = new Scanner(new File("marcas.txt"));
      String[] marcas = null;
      ArrayList data = new ArrayList();
      while (scan.hasNextLine()) {
        //d = new ArrayList();
        String line = scan.nextLine();
        //if(line.indexOf(",")>-1)
          //  array = line.split(",");
        //else
          //  array = line.split("\t");
        
        //for (int i = 0; i < array.length ; i ++){
          //  d.add(array[i]);
        //}
        //data.add(d);
        data.add(line);
        }
        //scan.close();
        marcas = new String[data.size()+1];
        marcas[0] = "---";
        for (int i = 0; i < data.size(); i++) 
            marcas[i+1] = data.get(i).toString();
        
        //System.out.println(Arrays.toString(marcas));
        return marcas;
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
    
    public static void saveData(int[] indexToCompare, JTable tabla, JComboBox marca, boolean se) throws Exception{
        Vector data = getDataFromTable(tabla);
        Vector newData = new Vector();
        Vector cells;
        for(int k = 0; k < data.size() ; k++){ //fila
         Vector row = (Vector) data.get(k);
         cells = new Vector();
         for(int i = 0 ; i < 14; i++)
             if(i == 13 )
              cells.add(marca.getSelectedItem().toString());
             else cells.add("");
          //for(int i = 0; i < row.size() ; i ++){//columna
              for(int j = 0; j < indexToCompare.length ; j++)
                  if(indexToCompare[j] != 0){
                     cells.setElementAt(row.get(j), indexToCompare[j]-1);
                     //System.out.println(cells.get(j)+"-"+row.get(j));
              }
              newData.add(cells); 
          }         
          
            //indexToCompare[i] = ((JComboBox)this.jPanelComboBox.getComponent(i)).getSelectedIndex();
            //System.out.println(((JComboBox)this.jPanelComboBox.getComponent(i)).getSelectedIndex());
        //data = getDataFromTable(tabla);
        //System.out.println("Salvando datos."+ data);
        new ExcelWordModel().saveExcel(newData, marca.getSelectedItem().toString(),marca.getSelectedIndex(),se);
        //sortVector(indexToCompare,data);
    }
    
    public static void saveData(int[] indexToCompare, JTable tabla, JComboBox marca, String sFileName, String sFilePath) throws Exception{
        Vector data = getDataFromTable(tabla);
        Vector newData = new Vector();
        Vector cells;
        for(int k = 0; k < data.size() ; k++){ //fila
         Vector row = (Vector) data.get(k);
         cells = new Vector();
         for(int i = 0 ; i < 14; i++)
             if(i == 13 )
              cells.add(marca.getSelectedItem().toString());
             else cells.add("");
          //for(int i = 0; i < row.size() ; i ++){//columna
              for(int j = 0; j < indexToCompare.length ; j++)
                  if(indexToCompare[j] != 0){
                     cells.setElementAt(row.get(j), indexToCompare[j]-1);
                     //System.out.println(cells.get(j)+"-"+row.get(j));
              }
              newData.add(cells); 
          }         
          
            //indexToCompare[i] = ((JComboBox)this.jPanelComboBox.getComponent(i)).getSelectedIndex();
            //System.out.println(((JComboBox)this.jPanelComboBox.getComponent(i)).getSelectedIndex());
        //data = getDataFromTable(tabla);
        //System.out.println("Salvando datos."+ data);
        new ExcelWordModel().saveExcel(newData, marca.getSelectedItem().toString(),marca.getSelectedIndex(), sFileName, sFilePath);
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
    
    public static void createCSV( String path ) throws ParseException{
        new Tabula(new String[]{path,"-r", "-p all", "-o " + ViewController.pathInicioUsuario});
    }
    
    public static String[] splitString(String line){
     //String line = "foo,bar,c;qual=\"baz,blurb\",d;junk=\"quux,syzygy\"";

        String otherThanQuote = " [^\"] ";
        String quotedString = String.format(" \" %s* \" ", otherThanQuote);
        String regex = String.format("(?x) "+ // enable comments, ignore white spaces
                ",                         "+ // match a comma
                "(?=                       "+ // start positive look ahead
                "  (                       "+ //   start group 1
                "    %s*                   "+ //     match 'otherThanQuote' zero or more times
                "    %s                    "+ //     match 'quotedString'
                "  )*                      "+ //   end group 1 and repeat it zero or more times
                "  %s*                     "+ //   match 'otherThanQuote'
                "  $                       "+ // match the end of the string
                ")                         ", // stop positive look ahead
                otherThanQuote, quotedString, otherThanQuote);

        String[] tokens = line.split(regex, -1);
        
        //for(String t : tokens) System.out.println("> "+t);
        
        return tokens;
    }
    
    public static CustomTableModel fillVectorCSV( String tabula ) {
        //Tabula tb = null;
        //String tabula = "";
        Vector<Vector> data = new Vector();
        Vector<String> row = null;
        int mayor = 0;
        String[] s2 = null;
        String[] s = tabula.split("\\r?\\n");
        //row = new Vector<String>();
        //System.err.print(s[0]+"");
        for (int i = 0; i < s.length; i++){
             s2 = splitString(s[i]);
             row = new Vector<String>();
          for (int j = 0; j < s2.length; j++){
            row.add(s2[j]);
            /*if(s[j].contains("\\n")){
                data.add(row);
                System.out.println();
                row = new Vector<String>();
            }*/
          }
          if(mayor < row.size())
              mayor = row.size();
          data.add(row);
          //System.out.print(mayor+"|");
          //System.out.println(data.size()+"-");
        }
        //System.out.println(Arrays.deepToString(data.toArray()));
        //ExcelWordModel.printVector(data);
        //CustomTableModel ctm = new CustomTableModel();
        //ctm.setColumnCount(mayor);
        //ctm.setRowCount(data.size());
        //System.out.println(Arrays.deepToString(data.toArray()));
        
        for (int i = 0 ; i < data.size() ; i++){
          for(int j = 0 ; j < mayor ;  j++){
             if(mayor > ((Vector)data.get(i)).size()){
              ((Vector)data.get(i)).add("");
              //ctm.setValueAt(((Vector)data.get(i)).get(j), i, j);
             } else{
               //ctm.setValueAt(((Vector)data.get(i)).get(j), i, j);
               //System.out.println(i+".- "+ctm.getValueAt(i, j));
              }
          }
        }
        Vector header = new Vector();
        for(int i = 0; i < row.size(); i++)
            header.add(indexColumn[i]);
        
        int i = 1;
        for(Vector t : data)
         System.out.println((i++)+".- colNum: "+t.size()+"-"+ Arrays.deepToString(t.toArray()) );
        //ctm.setDataVector(data, new Vector());
        return new CustomTableModel(data, header);
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
      //System.out.println(model.getRowCount()+" -- "+ model.getColumnCount() +": col = "+ col);
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
    
    public static Vector insertDataCSV(String source) throws FileNotFoundException, IOException {
        //CustomTableModel model = new CustomTableModel();
        CSVReader reader = new CSVReader(new FileReader(source));
        String [] Line;
        int col = 0;
        int filas = 0;
        Vector<Vector> data = new Vector();
        Vector<String> d = null;
        while ((Line = reader.readNext()) != null) {
             d = new Vector();
            // nextLine[] is an array of values from the line
            for(int i = 0; i < Line.length; i++){
                 d.add(Line[i]);
            }
            if(Line.length > col)
                col = Line.length;
            data.add(d);
        }
        //model.setColumnCount(col);
        //model.setRowCount(data.size());
        //System.out.println(model.getRowCount()+"-"+model.getColumnCount());
//        for(int i = 0; i < data.size() ; i++)
//            for(int j = 0; j< col; j++){
//                model.setValueAt(((Vector<String>)data.get(i)).get(j), i, j);
//                System.out.println("--"+model.getValueAt(i, j));//((Vector)data.get(i)).get(j).toString());
//            }
//        model.setDataVector(data, new Vector());
        
    return data;//model;
} 
    
    public static CustomTableModel fillTableCSV(Vector a){
        CustomTableModel dtm = new CustomTableModel();
        dtm.setColumnCount(((Vector)a.get(0)).size());
        dtm.setRowCount(a.size());
        dtm.setDataVector(a, new Vector());
        return dtm;
    }
    
    public static void main(String[] args){
        try {
            //ViewController.fillVectorCSV("C:\\Users\\Elpapo\\Downloads\\N-outline.pdf");
            ViewController.insertDataCSV("C:\\Users\\Elpapo\\Desktop\\extractoTablas.csv");
        } catch (IOException ex) {
            Logger.getLogger(ViewController.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    
}
