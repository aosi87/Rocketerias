/*
 * Copyright 2014 Elpapo.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package controller;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.ProcessBuilder.Redirect;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import org.apache.pdfbox.pdmodel.PDDocument;
//import org.jruby.embed.ScriptingContainer;
import org.nerdpower.tabula.ObjectExtractor;
import org.nerdpower.tabula.Page;
import org.nerdpower.tabula.PageIterator;
import org.nerdpower.tabula.Rectangle;
import org.nerdpower.tabula.Table;
import org.nerdpower.tabula.Utils;
import org.nerdpower.tabula.extractors.BasicExtractionAlgorithm;
import org.nerdpower.tabula.extractors.SpreadsheetExtractionAlgorithm;
import org.nerdpower.tabula.writers.CSVWriter;
import org.nerdpower.tabula.writers.JSONWriter;
import org.nerdpower.tabula.writers.TSVWriter;
import org.nerdpower.tabula.writers.Writer;

/**
 *
 * @author Isaac Alcocer
 */
public class TabulaController {
    
    public void checkProcess() throws Exception {
        //*
        ProcessBuilder pb = new ProcessBuilder(
                "java",
                "-jar",
                "C:\\Users\\Elpapo\\Downloads\\jruby-complete-1.7.13.jar",
                "\"bin/jruby -v\"");
        pb.redirectOutput(Redirect.INHERIT);
        pb.redirectError(Redirect.INHERIT);
        Process p = pb.start();
        p.waitFor();
        p.destroy();
        /*/
        
        /*
        Process p = Runtime.getRuntime().exec(
              "java -Xmx500m -Xss1024k -jar C:\\Users\\Elpapo\\Downloads\\jruby-complete-1.7.13.jar "
              + "-e \"puts 'Hello'\" ");
        p.waitFor();

        String line;

        BufferedReader error = new BufferedReader(new InputStreamReader(p.getErrorStream()));
        while((line = error.readLine()) != null){
            System.out.println(line);
        }
        error.close();

        BufferedReader input = new BufferedReader(new InputStreamReader(p.getInputStream()));
        while((line=input.readLine()) != null){
            System.out.println(line);
        }

        input.close();

        OutputStream outputStream = p.getOutputStream();
        PrintStream printStream = new PrintStream(outputStream);
        printStream.println();
        printStream.flush();
        printStream.close();*/
        
    }
    
    //Jsr223 version
    private TabulaController() throws ScriptException {
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine engine = manager.getEngineByName("jruby");
        engine.eval("puts 'Hello World!'");
    }
    
    /*embed version
    private TabulaController(String a){
       ScriptingContainer container = new ScriptingContainer();
       container.runScriptlet(a);
    }*/
    
    static void extractTables(String[] line) throws IOException {
        //El primer argumento es el archivo a leer.
        File pdfFile = new File(line[0]);
        if (!pdfFile.exists()) {
            throw new IOException("File does not exist");
        }
        //formato del archivo de salida.
        OutputFormat of = OutputFormat.CSV;
        //Archivo de salida, si es definido.
        Appendable outFile = System.out;
        if (line[1].length() > 1 && !line[1].isEmpty()) { //o
            File file = new File(line[1]);
            
            try {
                file.createNewFile();
                outFile = new BufferedWriter(new FileWriter(
                        file.getAbsoluteFile()));
            } catch (IOException e) {
                throw new IOException("Cannot create file "
                        + line[1]);
            }
        }
        
        Rectangle area = null;
        if (line.length > 2 && !line[2].isEmpty()) { //a
            List<Float> f = parseFloatList(line[2]);
            if (f.size() != 4) {
                throw new IOException("area parameters must be top,left,bottom,right");
            }
            area = new Rectangle(f.get(0), f.get(1), f.get(3) - f.get(1), f.get(2) - f.get(0));
        }
        
        List<Float> verticalRulingPositions = null;
        if (line.length > 3 && !line[3].isEmpty()) {//c
            verticalRulingPositions = parseFloatList(line[3]);
        }
        List<Integer> pages = parsePagesOption("all");
        if(line.length > 4 && !line[4].isEmpty())
            pages = parsePagesOption(line[4]);
        
        ExtractionMethod method = whichExtractionMethod(line[5]);
        
        boolean useLineReturns = false;
        if(line.length > 6 && !line[6].isEmpty())
            useLineReturns = true;
        
        boolean password = false;
        if(line.length > 7 && !line[7].isEmpty())
            password = true;
            
        try {
     
            ObjectExtractor oe = password ? 
                    new ObjectExtractor(PDDocument.load(pdfFile), line[7]) : 
                    new ObjectExtractor(PDDocument.load(pdfFile));
            BasicExtractionAlgorithm basicExtractor = new BasicExtractionAlgorithm();
            SpreadsheetExtractionAlgorithm spreadsheetExtractor = new SpreadsheetExtractionAlgorithm();
                    
            PageIterator pageIterator = pages == null ? oe.extract() : oe.extract(pages);
            Page page;
            List<Table> tables = new ArrayList<Table>();

            while (pageIterator.hasNext()) {
                page = pageIterator.next();
                
                if (area != null) {
                    page = page.getArea(area);
                }

                if (method == ExtractionMethod.DECIDE) {
                    method = spreadsheetExtractor.isTabular(page) ? ExtractionMethod.SPREADSHEET : ExtractionMethod.BASIC;
                }
                
                switch(method) {
                case BASIC:
                    if (line.length > 7 && !line[7].isEmpty()) {
                        
                    }
                    tables.addAll(verticalRulingPositions == null ? basicExtractor.extract(page) : basicExtractor.extract(page, verticalRulingPositions));
                    
                    break;
                case SPREADSHEET:
                    // TODO add useLineReturns
                    tables.addAll(spreadsheetExtractor.extract(page));
                default:
                    break;
                }
                for (Table t: tables) {
                    writeTables(of, tables, outFile);
                }
                tables.clear();
            }

        } catch (IOException e) {
            throw new IOException(e.getMessage());
        } 

    }
    
    static void writeTables(OutputFormat format, Iterable<Table> tables, Appendable out) throws IOException {
        Writer writer = null;
        switch (format) {
        case CSV:
            writer = new CSVWriter();
            break;
        case JSON:
            writer = new JSONWriter();
            break;
        case TSV:
            writer = new TSVWriter();
            break;
        }
        Iterator<Table> iter = tables.iterator();
        while (iter.hasNext()) {
            writer.write(out, iter.next());
        }
    }
    
    static ExtractionMethod whichExtractionMethod(String line) {
        ExtractionMethod rv = ExtractionMethod.DECIDE;
        if (line.equals('r')) {
            rv = ExtractionMethod.SPREADSHEET;
        }
        else if (line.equals('n') || line.equals('c') || line.equals('a') || line.equals('g')) {
            rv = ExtractionMethod.BASIC;
        }
        return rv;
    }
    
    public static List<Integer> parsePagesOption(String pagesSpec) throws IOException {
        if (pagesSpec == "all") {
            return null;
        }
        
        List<Integer> rv = new ArrayList<Integer>();
        
        String[] ranges = pagesSpec.split(",");
        for (int i = 0; i < ranges.length; i++) {
            String[] r = ranges[i].split("-");
            if (r.length == 0 || !Utils.isNumeric(r[0]) || r.length > 1 && !Utils.isNumeric(r[1])) {
                throw new IOException("Syntax error in page range specification");
            }
            
            if (r.length < 2) {
                rv.add(Integer.parseInt(r[0]));
            }
            else {
                int t = Integer.parseInt(r[0]);
                int f = Integer.parseInt(r[1]); 
                if (t > f) {
                    throw new IOException("Syntax error in page range specification");
                }
                rv.addAll(Utils.range(t, f+1));       
            }
        }
        
        Collections.sort(rv);
        return rv;
    }
    
    public static List<Float> parseFloatList(String option) throws IOException {
        String[] f = option.split(",");
        List<Float> rv = new ArrayList<Float>();
        try {
            for (int i = 0; i < f.length; i++) {
                rv.add(Float.parseFloat(f[i]));
            }
            return rv;
        } catch (NumberFormatException e) {
            throw new IOException("Wrong number syntax");
        }
    }
    
    
    private enum OutputFormat {
        CSV,
        TSV,
        JSON;
        
        static String[] formatNames() {
            OutputFormat[] values = OutputFormat.values();
            String[] rv = new String[values.length];
            for (int i = 0; i < values.length; i++) {
                rv[i] = values[i].name();
            }
            return rv;
        }
        
    }
        
    private enum ExtractionMethod {
        BASIC,
        SPREADSHEET,
        DECIDE
    }
    
    private class DebugOutput {
        private boolean debugEnabled;

        public DebugOutput(boolean debug) {
            this.debugEnabled = debug;
        }
        
        public void debug(String msg) {
            if (this.debugEnabled) {
                System.err.println(msg);    
            }
        }
    }

    public static void main(String[] args) throws ScriptException {
        String datos[] = new String[8];
        datos[0] = "C:/Users/Manuu Alcocer/Downloads/00Archivos Muestras/Avenview Master Cust Price List Q4 2012_REV_2.pdf"; //archivo a analizar
        datos[1] = ViewController.pathExcelSalidaDefault + "/CSV.csv";//archivo salida
        datos[2] = ""; //Rectangulo "float,float,float,float"
        datos[3] = ""; //coordenadas de columnas "float,float,float"
        datos[4] = ""; //Paginas por defecto es "all", de lo contrario "1" "3" "1-3" "5-7"
        datos[5] = ""; // "r" "n" "c" "a" "g"
        datos[6] = ""; // con ser diferente de una cadena vacia se toma como verdadero
        datos[7] = ""; // password del documento si es que tiene password
        
        try {    
            TabulaController.extractTables(datos);
        } catch (IOException ex) {
            Logger.getLogger(TabulaController.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }
    
    
    
}
