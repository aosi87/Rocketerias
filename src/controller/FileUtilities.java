/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package controller;

//JavaIO     
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;

//Collections
import java.util.ArrayList;
import java.util.Iterator;
// Actualice a ArrayList ya que tecnicamente Vector es obsoleto desde jdk 1.2
//import java.util.Vector; 

//Librerias para Apache PDFBox
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;

/* Librerias exclusivas de Apache POI para Microsoft Office Excel*/
//Excel 2010 en adelante (xlsx)
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//Excel 2007 y anteriores (xls)
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Utilerias:
 *  En esta clase se localizan las funciones necesarias para la obtencion y
 *  clasificacion de diversos datos, ademas de su uso en la aplicacion se 
 *  plantean metodos para la administracion basica, buscar, borrar, agregar,
 *  unir, crear, separar, etc.
 *  
 * La funcionalidad basica recibe la ruta del archivo a trabajar, verificar que exista,
 * asi como identificar el tipo de archivo (PDF, XLSX, DOCX), para despues trabajar con
 * el mismo.
 * 
 * Casos de uso:
 * 1.- FileUtilities fu = new FileUtilities("ruta/nombre.extencion");
 *  int numHojas = fu.obtenerHojas();
 * 
 * 2.- Iterator column = getColumns(new FileUtilities("ruta/nombre.extencion"));
 * 
 * 
 * @author Isaac Alcocer.
 */
public class FileUtilities {

    public String path;
    XSSFWorkbook wb = null;
    
    public FileUtilities(String path){
        this.path = path;
    }
    
}//fin
