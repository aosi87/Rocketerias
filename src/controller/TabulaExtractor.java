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

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.cli.ParseException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.nerdpower.tabula.ObjectExtractor;
import org.nerdpower.tabula.Page;
import org.nerdpower.tabula.PageIterator;
import org.nerdpower.tabula.Rectangle;
import org.nerdpower.tabula.RectangularTextContainer;
import org.nerdpower.tabula.Table;
import org.nerdpower.tabula.extractors.BasicExtractionAlgorithm;
import org.nerdpower.tabula.extractors.SpreadsheetExtractionAlgorithm;
import org.nerdpower.tabula.writers.CSVWriter;
import org.nerdpower.tabula.writers.JSONWriter;
import org.nerdpower.tabula.writers.TSVWriter;
import org.nerdpower.tabula.writers.Writer;

/**
 *
 * @author Elpapo
 */
public class TabulaExtractor {
    private boolean GUESS;
    
    private enum ExtractionMethod {
        BASIC,
        SPREADSHEET,
        DECIDE
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
    
    String path = null;
    SpreadsheetExtractionAlgorithm sea = null;
    BasicExtractionAlgorithm bea = null;
    
    public TabulaExtractor(String path){
        this.path = path;
    }
    
    private Table getTableBasic(int index) throws IOException {
        Page page = this.getPage(index);//UtilsForTesting.getAreaFromFirstPage("src/test/resources/org/nerdpower/tabula/argentina_diputados_voting_record.pdf", 269.875f, 12.75f, 790.5f, 561f);
        bea = new BasicExtractionAlgorithm();
        Table table = bea.extract(page).get(0);
        return table;
    }
    
    private Table getTableSpreadSheet(int index) throws IOException {
        Page page = this.getPage(index);//UtilsForTesting.getAreaFromFirstPage("src/test/resources/org/nerdpower/tabula/argentina_diputados_voting_record.pdf", 269.875f, 12.75f, 790.5f, 561f);
        sea = new SpreadsheetExtractionAlgorithm();
        Table table = sea.extract(page).get(index);
        return table;
    }
    
    public PageIterator getPages() throws IOException {
        ObjectExtractor oe = null;
        try {
            PDDocument document = PDDocument
                    .load(path);
            oe = new ObjectExtractor(document);
            PageIterator pageIter = oe.extract();
            return pageIter;
        } finally {
            oe.close();
        }
    }
    
    public Page getPage(int index) throws IOException {
        ObjectExtractor oe = null;
        try {
            PDDocument document = PDDocument
                    .load(path);
            oe = new ObjectExtractor(document);
            Page page = oe.extract(index);
            return page;
        } finally {
            oe.close();
        }
    }
    
    public void basicExtractor() throws IOException{
        //Table table = this.getTableSpreadSheet(1);
        Table table = this.getTableBasic(1);
        StringBuilder sb = new StringBuilder();
        (new CSVWriter()).write(sb, table);
        String s = sb.toString();
        String[] lines = s.split("\\r?\\n");
        System.err.print(s);
    }
    
    public void smartExtractor(PageIterator pageIterator) throws IOException, ParseException{
        Page page = null;
        Rectangle area = null;
        ExtractionMethod method = ExtractionMethod.BASIC;
        List<Table> tables = new ArrayList<>();
        List<Float> verticalRulingPositions = null;
        
        while (pageIterator.hasNext()) {
            page = pageIterator.next();
            System.err.print(pageIterator.next());
            
            if (area != null) {
                page = page.getArea(area);
            }
            
            if (method == ExtractionMethod.DECIDE) {
                method = sea.isTabular(page) ? ExtractionMethod.SPREADSHEET : ExtractionMethod.BASIC;
            }
            
            switch(method) {
                case BASIC:
                    if (GUESS == true) {
                        verticalRulingPositions = parseFloatList("0,0,0,0");
                    }
                    tables.addAll(verticalRulingPositions == null ? bea.extract(page) : bea.extract(page, verticalRulingPositions));
                    
                    break;
                case SPREADSHEET:
                    // TODO add useLineReturns
                    tables.addAll(sea.extract(page));
                default:
                    break;
            }
            for (Table t: tables) {
                writeVector(t);
            }
            tables.clear();
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
    
    public static List<Float> parseFloatList(String option) throws ParseException {
        String[] f = option.split(",");
        List<Float> rv = new ArrayList<Float>();
        try {
            for (int i = 0; i < f.length; i++) {
                rv.add(Float.parseFloat(f[i]));
            }
            return rv;
        } catch (NumberFormatException e) {
            throw new ParseException("Wrong number syntax");
        }
    }
    
    public Vector writeVector(Table table) throws IOException {
        Vector data = new Vector();
        for (List<RectangularTextContainer> row: table.getRows()) {
            List<String> cells = new ArrayList<String>(row.size());
            for (RectangularTextContainer tc: row) {
                cells.add(tc.getText());
            }
            data.add(cells);
        }
        System.out.println(data);
        return data;
    }        
    
    
    public static void main(String args[]){
        TabulaExtractor te = new TabulaExtractor("C:\\Users\\Elpapo\\Desktop\\Price List 2014 - Confidential.pdf");
        try {
            te.basicExtractor();//smartExtractor(te.getPages());
        } catch (IOException ex) {
            Logger.getLogger(TabulaExtractor.class.getName()).log(Level.SEVERE, null, ex);
        }
    } 
}
