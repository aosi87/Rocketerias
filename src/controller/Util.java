/*
 * Copyright 2014 Isaac Alcocer.
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

import java.io.File;  
import java.io.FileNotFoundException;
import java.io.FileOutputStream;  
import java.io.IOException;
import java.util.ArrayList;  
import java.util.HashMap;  
import java.util.Iterator;  
import java.util.List;  
import java.util.Map;  
import java.util.Set;  
import java.util.TreeSet;  
import java.util.logging.Level;
import java.util.logging.Logger;
import model.ExcelWordModel;  
import org.apache.poi.hssf.usermodel.HSSFCell;  
import org.apache.poi.hssf.usermodel.HSSFCellStyle;  
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;  
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.CellStyle;  
import org.apache.poi.ss.usermodel.ClientAnchor;  
import org.apache.poi.ss.usermodel.CreationHelper;  
import org.apache.poi.ss.usermodel.DataFormat;  
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
          
      /**   
       *   
       * @author jk   
       * getted from http://jxls.cvs.sourceforge.net/jxls/jxls/src/java/org/jxls/util/Util.java?revision=1.8&view=markup   
       * by Leonid Vysochyn    
       * and modified (adding styles copying)   
       * modified by Philipp Löpmeier (replacing deprecated classes and methods, using generic types)   
       */    
      public final class Util {     
               
            /** 
             * DEFAULT CONSTRUCTOR. 
             */  
            private Util() {}  
              
          /** 
           * @param newSheet the sheet to create from the copy. 
           * @param sheet the sheet to copy. 
           */  
          public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet){     
              copySheets(newSheet, sheet, true);     
          }     
               
          /** 
           * @param newSheet the sheet to create from the copy. 
           * @param sheet the sheet to copy. 
           * @param copyStyle true copy the style. 
           */  
          public static void copySheets(XSSFSheet newSheet, XSSFSheet sheet, boolean copyStyle){     
              int maxColumnNum = 0;     
              Map<Integer, XSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, XSSFCellStyle>() : null;     
              for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {     
                  XSSFRow srcRow = sheet.getRow(i);     
                  XSSFRow destRow = newSheet.createRow(i);     
                  if (srcRow != null) {     
                      Util.copyRow(sheet, newSheet, srcRow, destRow, styleMap);     
                      if (srcRow.getLastCellNum() > maxColumnNum) {     
                          maxColumnNum = srcRow.getLastCellNum();     
                      }     
                  }     
              }     
              for (int i = 0; i <= maxColumnNum; i++) {     
                  newSheet.setColumnWidth(i, sheet.getColumnWidth(i));     
              }
              //Util.copyPictures(newSheet, sheet);
          }     
          
          /** 
           * @param srcSheet the sheet to copy. 
           * @param destSheet the sheet to create. 
           * @param srcRow the row to copy. 
           * @param destRow the row to create. 
           * @param styleMap - 
           */  
          public static void copyRow(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFRow srcRow, XSSFRow destRow, Map<Integer, XSSFCellStyle> styleMap) {     
              // manage a list of merged zone in order to not insert two times a merged zone  
              List list = new ArrayList();
              Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();     
              destRow.setHeight(srcRow.getHeight());     
              // pour chaque row  
              int j = srcRow.getFirstCellNum(); 
              int deltaRow = destRow.getRowNum()-srcRow.getRowNum();
              if(j<0){j=0;} 
                for (; j <= srcRow.getLastCellNum(); j++) {
              //for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {     
                  XSSFCell oldCell = srcRow.getCell(j);   // ancienne cell  
                  XSSFCell newCell = destRow.getCell(j);  // new cell   
                  if (oldCell != null) {     
                      if (newCell == null) {     
                          newCell = destRow.createCell(j);     
                      }     
                      // copy chaque cell  
                      //copyCell(oldCell, newCell, styleMap); 
                      for (int i = 0; i < styleMap.size(); i++)
                       list.add(styleMap.get(j));
                      copyCell((Cell)oldCell, (Cell)newCell, list);
                      // copy les informations de fusion entre les cellules  
                      //System.out.println("row num: " + srcRow.getRowNum() + " , col: " + (short)oldCell.getColumnIndex());  
                      CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(), (short)oldCell.getColumnIndex());     
                        
                      if (mergedRegion != null) {   
                        //System.out.println("Selected merged region: " + mergedRegion.toString());  
                        CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow()+deltaRow, mergedRegion.getLastRow()+deltaRow, mergedRegion.getFirstColumn(),  mergedRegion.getLastColumn());  
                          //System.out.println("New merged region: " + newMergedRegion.toString());  
                          CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);  
                          if (isNewMergedRegion(wrapper, mergedRegions)) {  
                              mergedRegions.add(wrapper);  
                              destSheet.addMergedRegion(wrapper.range);     
                          }     
                      }     
                  }     
              }     
                   
          }     
               
          /** 
           * @param oldCell 
           * @param newCell 
           * @param styleMap 
           */  
          public static void copyCell(XSSFCell oldCell, XSSFCell newCell, Map<Integer, XSSFCellStyle> styleMap) {     
              if(styleMap != null) {     
                  if(oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()){     
                      newCell.setCellStyle(oldCell.getCellStyle());     
                  } else{     
                      int stHashCode = oldCell.getCellStyle().hashCode();     
                      XSSFCellStyle newCellStyle = styleMap.get(stHashCode);     
                      if(newCellStyle == null){     
                          newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();     
                          newCellStyle.cloneStyleFrom(oldCell.getCellStyle());     
                          styleMap.put(stHashCode, newCellStyle);     
                      }     
                      newCell.setCellStyle(newCellStyle);     
                  }     
              }     
              switch(oldCell.getCellType()) {     
                  case XSSFCell.CELL_TYPE_STRING:     
                      newCell.setCellValue(oldCell.getStringCellValue());     
                      break;     
                  case XSSFCell.CELL_TYPE_NUMERIC:     
                      newCell.setCellValue(oldCell.getNumericCellValue());     
                      break;     
                  case XSSFCell.CELL_TYPE_BLANK:     
                      newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);     
                      break;     
                  case XSSFCell.CELL_TYPE_BOOLEAN:     
                      newCell.setCellValue(oldCell.getBooleanCellValue());     
                      break;     
                  case XSSFCell.CELL_TYPE_ERROR:     
                      newCell.setCellErrorValue(oldCell.getErrorCellValue());     
                      break;     
                  case XSSFCell.CELL_TYPE_FORMULA:     
                      newCell.setCellFormula(oldCell.getCellFormula());     
                      break;     
                  default:     
                      break;     
              }     
                   
          }    
          
    public static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {  
        if (styleList != null) {  
            if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {  
            newCell.setCellStyle(oldCell.getCellStyle());  
        } else {  
           DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();  
  
        CellStyle newCellStyle = getSameCellStyle(oldCell, newCell, styleList);  
        if (newCellStyle == null) {  
            //Create a new cell style  
            Font oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());  
            //Find a existing font corresponding to avoid to create a new one   
            Font newFont = newCell.getSheet().getWorkbook().findFont(oldFont.getBoldweight(), oldFont.getColor(), oldFont.getFontHeight(), oldFont.getFontName(), oldFont.getItalic(), oldFont.getStrikeout(), oldFont.getTypeOffset(), oldFont.getUnderline());  
            if (newFont == null) {  
            newFont = newCell.getSheet().getWorkbook().createFont();  
            newFont.setBoldweight(oldFont.getBoldweight());  
            newFont.setColor(oldFont.getColor());  
            newFont.setFontHeight(oldFont.getFontHeight());  
            newFont.setFontName(oldFont.getFontName());  
            newFont.setItalic(oldFont.getItalic());  
            newFont.setStrikeout(oldFont.getStrikeout());  
            newFont.setTypeOffset(oldFont.getTypeOffset());  
            newFont.setUnderline(oldFont.getUnderline());  
            newFont.setCharSet(oldFont.getCharSet());  
            }  
  
            short newFormat = newDataFormat.getFormat(oldCell.getCellStyle().getDataFormatString());  
            newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();  
            newCellStyle.setFont(newFont);  
            newCellStyle.setDataFormat(newFormat);  
  
            newCellStyle.setAlignment(oldCell.getCellStyle().getAlignment());  
            newCellStyle.setHidden(oldCell.getCellStyle().getHidden());  
            newCellStyle.setLocked(oldCell.getCellStyle().getLocked());  
            newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());  
            newCellStyle.setBorderBottom(oldCell.getCellStyle().getBorderBottom());  
            newCellStyle.setBorderLeft(oldCell.getCellStyle().getBorderLeft());  
            newCellStyle.setBorderRight(oldCell.getCellStyle().getBorderRight());  
            newCellStyle.setBorderTop(oldCell.getCellStyle().getBorderTop());  
            newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());  
            newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());  
            newCellStyle.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundColor());  
            newCellStyle.setFillPattern(oldCell.getCellStyle().getFillPattern());  
            newCellStyle.setIndention(oldCell.getCellStyle().getIndention());  
            newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());  
            newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());  
            newCellStyle.setRotation(oldCell.getCellStyle().getRotation());  
            newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());  
            newCellStyle.setVerticalAlignment(oldCell.getCellStyle().getVerticalAlignment());  
  
            styleList.add(newCellStyle);  
        }  
        newCell.setCellStyle(newCellStyle);  
        }  
    }  
    switch (oldCell.getCellType()) {  
    case Cell.CELL_TYPE_STRING:  
        newCell.setCellValue(oldCell.getStringCellValue());  
        break;  
    case Cell.CELL_TYPE_NUMERIC:  
        newCell.setCellValue(oldCell.getNumericCellValue());  
        break;  
    case Cell.CELL_TYPE_BLANK:  
        newCell.setCellType(Cell.CELL_TYPE_BLANK);  
        break;  
    case Cell.CELL_TYPE_BOOLEAN:  
        newCell.setCellValue(oldCell.getBooleanCellValue());  
        break;  
    case Cell.CELL_TYPE_ERROR:  
        newCell.setCellErrorValue(oldCell.getErrorCellValue());  
        break;  
    case Cell.CELL_TYPE_FORMULA:  
        newCell.setCellFormula(oldCell.getCellFormula());  
        break;  
    default:  
        break;  
    }  
  
    }  
    
    private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {  
    CellStyle styleToFind = oldCell.getCellStyle();  
    CellStyle currentCellStyle = null;  
    CellStyle returnCellStyle = null;  
    Iterator<CellStyle> iterator = styleList.iterator();  
    Font oldFont = null;  
    Font newFont = null;  
    while (iterator.hasNext() && returnCellStyle == null) {  
        currentCellStyle = iterator.next();  
  
        if (currentCellStyle.getAlignment() != styleToFind.getAlignment()) {  
        continue;  
        }  
        if (currentCellStyle.getHidden() != styleToFind.getHidden()) {  
        continue;  
        }  
        if (currentCellStyle.getLocked() != styleToFind.getLocked()) {  
        continue;  
        }  
        if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {  
        continue;  
        }  
        if (currentCellStyle.getBorderBottom() != styleToFind.getBorderBottom()) {  
        continue;  
        }  
        if (currentCellStyle.getBorderLeft() != styleToFind.getBorderLeft()) {  
        continue;  
        }  
        if (currentCellStyle.getBorderRight() != styleToFind.getBorderRight()) {  
        continue;  
        }  
        if (currentCellStyle.getBorderTop() != styleToFind.getBorderTop()) {  
        continue;  
        }  
        if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {  
        continue;  
        }  
        if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {  
        continue;  
        }  
        if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {  
        continue;  
        }  
        if (currentCellStyle.getFillPattern() != styleToFind.getFillPattern()) {  
        continue;  
        }  
        if (currentCellStyle.getIndention() != styleToFind.getIndention()) {  
        continue;  
        }  
        if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {  
        continue;  
        }  
        if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {  
        continue;  
        }  
        if (currentCellStyle.getRotation() != styleToFind.getRotation()) {  
        continue;  
        }  
        if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {  
        continue;  
        }  
        if (currentCellStyle.getVerticalAlignment() != styleToFind.getVerticalAlignment()) {  
        continue;  
        }  
  
        oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());  
        newFont = newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());  
  
        if (newFont.getBoldweight() == oldFont.getBoldweight()) {  
        continue;  
        }  
        if (newFont.getColor() == oldFont.getColor()) {  
        continue;  
        }  
        if (newFont.getFontHeight() == oldFont.getFontHeight()) {  
        continue;  
        }  
        if (newFont.getFontName() == oldFont.getFontName()) {  
        continue;  
        }  
        if (newFont.getItalic() == oldFont.getItalic()) {  
        continue;  
        }  
        if (newFont.getStrikeout() == oldFont.getStrikeout()) {  
        continue;  
        }  
        if (newFont.getTypeOffset() == oldFont.getTypeOffset()) {  
        continue;  
        }  
        if (newFont.getUnderline() == oldFont.getUnderline()) {  
        continue;  
        }  
        if (newFont.getCharSet() == oldFont.getCharSet()) {  
        continue;  
        }  
        if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {  
        continue;  
        }  
  
        returnCellStyle = currentCellStyle;  
    }  
    return returnCellStyle;  
    }  
          
    private static void copyPictures(Sheet newSheet, Sheet sheet) {  
        Drawing drawingOld = sheet.createDrawingPatriarch();  
        Drawing drawingNew = newSheet.createDrawingPatriarch();  
        CreationHelper helper = newSheet.getWorkbook().getCreationHelper();  
  
        if (drawingNew instanceof HSSFPatriarch) {  
            List<HSSFShape> shapes = ((HSSFPatriarch) drawingOld).getChildren();  
            for (int i = 0; i < shapes.size(); i++) {  
             if (shapes.get(i) instanceof HSSFPicture) {  
                HSSFPicture pic = (HSSFPicture) shapes.get(i);  
                HSSFPictureData picdata = pic.getPictureData();  
                int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(), picdata.getFormat());  
                ClientAnchor anchor = null;  
             if (pic.getAnchor() != null) {  
                anchor = helper.createClientAnchor();  
                anchor.setDx1(((HSSFClientAnchor) pic.getAnchor()).getDx1());  
                anchor.setDx2(((HSSFClientAnchor) pic.getAnchor()).getDx2());  
                anchor.setDy1(((HSSFClientAnchor) pic.getAnchor()).getDy1());  
                anchor.setDy2(((HSSFClientAnchor) pic.getAnchor()).getDy2());  
                anchor.setCol1(((HSSFClientAnchor) pic.getAnchor()).getCol1());  
                anchor.setCol2(((HSSFClientAnchor) pic.getAnchor()).getCol2());  
                anchor.setRow1(((HSSFClientAnchor) pic.getAnchor()).getRow1());  
                anchor.setRow2(((HSSFClientAnchor) pic.getAnchor()).getRow2());  
                anchor.setAnchorType(((HSSFClientAnchor) pic.getAnchor()).getAnchorType());  
             }  
             drawingNew.createPicture(anchor, pictureIndex);  
             }  
            }  
        } else {  
            if (drawingNew instanceof XSSFDrawing) {  
            List<XSSFShape> shapes = ((XSSFDrawing) drawingOld).getShapes();  
            for (int i = 0; i < shapes.size(); i++) {  
             if (shapes.get(i) instanceof XSSFPicture) {  
                XSSFPicture pic = (XSSFPicture) shapes.get(i);  
                XSSFPictureData picdata = pic.getPictureData();  
                int pictureIndex = newSheet.getWorkbook().addPicture(picdata.getData(), picdata.getPictureType());  
                XSSFClientAnchor anchor = null;  
                CTTwoCellAnchor oldAnchor = ((XSSFDrawing) drawingOld).getCTDrawing().getTwoCellAnchorArray(i);  
             if (oldAnchor != null) {  
                anchor = (XSSFClientAnchor) helper.createClientAnchor();  
                CTMarker markerFrom = oldAnchor.getFrom();  
                CTMarker markerTo = oldAnchor.getTo();  
                anchor.setDx1((int) markerFrom.getColOff());  
                anchor.setDx2((int) markerTo.getColOff());  
                anchor.setDy1((int) markerFrom.getRowOff());  
                anchor.setDy2((int) markerTo.getRowOff());  
                anchor.setCol1(markerFrom.getCol());  
                anchor.setCol2(markerTo.getCol());  
                anchor.setRow1(markerFrom.getRow());  
                anchor.setRow2(markerTo.getRow());  
             }  
             drawingNew.createPicture(anchor, pictureIndex);  
             }  
           }  
        }  
      }  
    }  
               
          /** 
           * Récupère les informations de fusion des cellules dans la sheet source pour les appliquer 
           * à la sheet destination... 
           * Récupère toutes les zones merged dans la sheet source et regarde pour chacune d'elle si 
           * elle se trouve dans la current row que nous traitons. 
           * Si oui, retourne l'objet CellRangeAddress. 
           *  
           * @param sheet the sheet containing the data. 
           * @param rowNum the num of the row to copy. 
           * @param cellNum the num of the cell to copy. 
           * @return the CellRangeAddress created. 
           */  
          public static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {     
              for (int i = 0; i < sheet.getNumMergedRegions(); i++) {   
                  CellRangeAddress merged = sheet.getMergedRegion(i);     
                  if (merged.isInRange(rowNum, cellNum)) {     
                      return merged;     
                  }     
              }     
              return null;     
          }     
          
          /** 
           * Check that the merged region has been created in the destination sheet. 
           * @param newMergedRegion the merged region to copy or not in the destination sheet. 
           * @param mergedRegions the list containing all the merged region. 
           * @return true if the merged region is already in the list or not. 
           */  
          private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion, Set<CellRangeAddressWrapper> mergedRegions) {  
            return !mergedRegions.contains(newMergedRegion);     
          } 
          
          public static void main(String args[]){
              FileOutputStream fos = null;
                try {
                    ExcelWordModel ewm = new ExcelWordModel(new File(ViewController.pathExcelSalidaDefault+"/Roketerias.xlsx"));
                    ExcelWordModel ewm2 = new ExcelWordModel(new XSSFWorkbook());
                    ewm2.getXLSXWorkbook().createSheet("Works");
                    Util.copySheets(ewm2.getXLSXWorkbook().getSheetAt(0), ewm.getXLSXWorkbook().getSheetAt(0));
                    fos = new FileOutputStream("C:\\Users\\Elpapo\\Desktop\\Demo.xlsx");
                    ewm2.getXLSXWorkbook().write(fos);
                    fos.close();
                } catch (FileNotFoundException ex) {
                    Logger.getLogger(Util.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    Logger.getLogger(Util.class.getName()).log(Level.SEVERE, null, ex);
                }
          }
               
      }   


class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {  
   
      public CellRangeAddress range;  
        
      /** 
       * @param theRange the CellRangeAddress object to wrap. 
       */  
      public CellRangeAddressWrapper(CellRangeAddress theRange) {  
            this.range = theRange;  
      }  
        
      /** 
       * @param o the object to compare. 
       * @return -1 the current instance is prior to the object in parameter, 0: equal, 1: after... 
       */  
      public int compareTo(CellRangeAddressWrapper o) {  
   
                  if (range.getFirstColumn() < o.range.getFirstColumn()  
                              && range.getFirstRow() < o.range.getFirstRow()) {  
                        return -1;  
                  } else if (range.getFirstColumn() == o.range.getFirstColumn()  
                              && range.getFirstRow() == o.range.getFirstRow()) {  
                        return 0;  
                  } else {  
                        return 1;  
                  }  
   
      }  
        
}  