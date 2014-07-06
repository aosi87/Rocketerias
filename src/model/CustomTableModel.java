/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package model;

import java.util.Vector;
import javax.swing.table.DefaultTableModel;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class CustomTableModel extends DefaultTableModel{
    private boolean isCellEditable;
    
    @SuppressWarnings("UseOfObsoleteCollectionType")
    public CustomTableModel( Vector vector, Vector vector1){
        super(vector, vector1);
        isCellEditable = false;
    }

    public CustomTableModel() {
        isCellEditable = false;
    }

    @Override
   public boolean isCellEditable(int row, int column) {
       //Only the third column
       return this.isCellEditable;
   }

    /**
     * @param isCellEditable the isCellEditable to set
     */
    public void setIsCellEditable(boolean isCellEditable) {
        this.isCellEditable = isCellEditable;
    }
    /*
    @Override
    public Class<?> getColumnClass(int columnIndex) {
             return this.getValueAt(0,columnIndex).getClass();
        }
    */
    
    
}
