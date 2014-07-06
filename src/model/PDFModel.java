/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package model;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class PDFModel {

    public PDFModel(File selectedFile) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
    
     public PDFModel(){}
    
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
        
        public static void main(String[] args){
            PDFModel pdf = new PDFModel();
            pdf.readPDF();
        }
       
    
}
