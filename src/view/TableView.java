/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package view;

import controller.ViewController;
import java.awt.Dialog;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.security.AccessControlException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Arrays;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFormattedTextField;
import javax.swing.JOptionPane;
import javax.swing.JTable;
import javax.swing.SwingWorker;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.text.NumberFormatter;
import model.CustomTableModel;
import org.apache.commons.cli.ParseException;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class TableView extends javax.swing.JFrame {
    
    private boolean bot1;
    private String path;
    private float factores[] = null;
    private int colIndex[];
    /*private String marcas[] = {"---","Starlighting","Arthea","OTIC Audio/Audio Logic",
                               "OTIC Led/Video Logic","Stealth Acoustic","Naim","Gefen",
                               "CurrentAudio","Fenix"};
    
    /**
     * Creates new form TableView
     */
    public TableView() {
        initComponents();
        this.setIconImage(Toolkit.getDefaultToolkit().getImage("resources/gfx/icon.png"));
        //this.jComboBox1 = new WideComboBox();
        this.jComboBox1.setModel(new DefaultComboBoxModel(ViewController.marcas));//this.marcas));
        //BoundsPopupMenuListener listener = new BoundsPopupMenuListener(true, false);
        //jComboBox1.addPopupMenuListener( listener );
        jComboBox1.setPrototypeDisplayValue("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
        System.out.println("entra al tableview");

    }
    
    public TableView(String path) {
        initComponents();
        this.jCheckBox1.setVisible(false);
        this.setIconImage(Toolkit.getDefaultToolkit().getImage("resources/gfx/icon.png"));
        //this.jTabbedPane1.setTitleAt(0, path);
        this.jComboBox1.setModel(new DefaultComboBoxModel(ViewController.marcas));//this.marcas));
        this.path = path;
        this.setLocationRelativeTo(null);
        this.jPanelManejoTabla.setBorder(null);
        this.jTable1.setRowSelectionAllowed(true);
        this.jTable1.setRowSelectionInterval(0, 0);
        this.bot1 = true;
        this.jTabbedPane1.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            path,
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
        this.jPanelManejoDatos.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            "Manejo de Datos",
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
        this.jPanelManejoTabla.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            "Manejo de Tabla",
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
        this.jPanelCombo.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            "Marcas",
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
    }
    
    void setIndexes(int[] indexes){
        this.colIndex = indexes;
    }
    
    
    void setSheetName(String s) {
        this.jTabbedPane1.setTitleAt(0, s);
    }
    
    public void setTableModel(DefaultTableModel model){
        this.jTable1.setModel(model);
        //this.jTable1.setValueAt(true, 0, 0);
        for(int i = 0;  i < this.jTable1.getColumnCount(); i++)
            this.jTable1.getColumnModel().getColumn(i).setMinWidth(100);
        this.repaint();
    }
    
    private boolean isARowSelected() {
        for (int i=0; i< this.jTable1.getRowCount(); i++)
            if(this.jTable1.isRowSelected(i))
                return true;
        return false;
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTabbedPane1 = new javax.swing.JTabbedPane();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jPanelManejoTabla = new javax.swing.JPanel();
        jButtonFilas = new javax.swing.JButton();
        jButtonCabeceras = new javax.swing.JButton();
        jButton1 = new javax.swing.JButton();
        jPanelManejoDatos = new javax.swing.JPanel();
        jButtonSeleccionColumn = new javax.swing.JButton();
        jButton7 = new javax.swing.JButton();
        jPanelCombo = new javax.swing.JPanel();
        jCheckBox1 = new javax.swing.JCheckBox();
        jComboBox1 = new javax.swing.JComboBox();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuOpen = new javax.swing.JMenuItem();
        jSeparator1 = new javax.swing.JPopupMenu.Separator();
        jMenuMarcas = new javax.swing.JMenuItem();
        jMenuItem3 = new javax.swing.JMenuItem();
        jSeparator2 = new javax.swing.JPopupMenu.Separator();
        jMenuItemExit = new javax.swing.JMenuItem();
        jMenu2 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Software Mineria de Datos.");
        setMinimumSize(new java.awt.Dimension(800, 600));
        setName("TableView"); // NOI18N

        jTabbedPane1.setMaximumSize(new java.awt.Dimension(87, 37));
        jTabbedPane1.setMinimumSize(new java.awt.Dimension(87, 37));
        jTabbedPane1.setPreferredSize(new java.awt.Dimension(87, 37));

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jTable1.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jTable1.setRowSelectionAllowed(false);
        jScrollPane1.setViewportView(jTable1);

        jTabbedPane1.addTab("Archivo", jScrollPane1);

        jPanelManejoTabla.setLayout(new java.awt.GridLayout(1, 0));

        jButtonFilas.setText("<html><center>Seleccionando:<br>por Fila</center></html>");
        jButtonFilas.setMaximumSize(new java.awt.Dimension(87, 37));
        jButtonFilas.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonFilasActionPerformed(evt);
            }
        });
        jPanelManejoTabla.add(jButtonFilas);

        jButtonCabeceras.setText("<html><center>Seleccionar <br>Cabeceras</center></html>");
        jButtonCabeceras.setMaximumSize(new java.awt.Dimension(87, 37));
        jButtonCabeceras.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonCabecerasActionPerformed(evt);
            }
        });
        jPanelManejoTabla.add(jButtonCabeceras);

        jButton1.setText("<html><center>Borrar<br>Seleccinados</center></html>");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        jPanelManejoTabla.add(jButton1);

        jPanelManejoDatos.setLayout(new java.awt.GridLayout(1, 0));

        jButtonSeleccionColumn.setText("<html><center>Escoger<br>Columnas</center></html>");
        jButtonSeleccionColumn.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonSeleccionColumnActionPerformed(evt);
            }
        });
        jPanelManejoDatos.add(jButtonSeleccionColumn);

        jButton7.setText("<html><center>Adjuntar<br>Datos</center></html>");
        jButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton7ActionPerformed(evt);
            }
        });
        jPanelManejoDatos.add(jButton7);

        jPanelCombo.setLayout(new java.awt.GridLayout(2, 0));

        jCheckBox1.setText("Guardar archivo nuevo");
        jCheckBox1.setBorder(javax.swing.BorderFactory.createEtchedBorder());
        jCheckBox1.setBorderPainted(true);
        jPanelCombo.add(jCheckBox1);

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "Marcas", "Item 2", "Item 3", "Item 4" }));
        jPanelCombo.add(jComboBox1);

        jMenu1.setText("Archivo");

        jMenuOpen.setText("Abrir");
        jMenuOpen.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuOpenActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuOpen);
        jMenu1.add(jSeparator1);

        jMenuMarcas.setText("Cargar Marcas");
        jMenu1.add(jMenuMarcas);

        jMenuItem3.setText("Guardar como...");
        jMenuItem3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem3ActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItem3);
        jMenu1.add(jSeparator2);

        jMenuItemExit.setText("Salir");
        jMenuItemExit.setToolTipText("Al seleccionar esta opcion el programa terminara todo proceso ejecutado por este programa,\nse perdera toda operacion no salvada con aterioridad y el programa finalizara y cerrara todas\nlas instancias del mismo.");
        jMenuItemExit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItemExitActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuItemExit);

        jMenuBar1.add(jMenu1);

        jMenu2.setText("Ayuda");

        jMenuItem1.setText("Acerca de");
        jMenuItem1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuItem1ActionPerformed(evt);
            }
        });
        jMenu2.add(jMenuItem1);

        jMenuBar1.add(jMenu2);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanelManejoTabla, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jPanelManejoDatos, javax.swing.GroupLayout.PREFERRED_SIZE, 220, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jPanelCombo, javax.swing.GroupLayout.PREFERRED_SIZE, 163, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
            .addComponent(jTabbedPane1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jPanelManejoDatos, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanelCombo, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanelManejoTabla, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 481, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jMenuItemExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItemExitActionPerformed
        // TODO add your handling code here:
        this.dispose();
    }//GEN-LAST:event_jMenuItemExitActionPerformed

    private void jMenuOpenActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuOpenActionPerformed
        // TODO add your handling code here:
        final JFileChooser fileChooser = new JFileChooser();
        //Agregamos un filtro de extensiones
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS-Word", "docx","doc"));
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Adobe-PDF", "pdf"));
        fileChooser.setFileFilter(new FileNameExtensionFilter("MS-Excel", "xlsx","xls"));
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        
        int returnValue = fileChooser.showDialog(this,"Seleccionar documento");
        if(fileChooser.getSelectedFile() != null){
            ViewController.setPDFCSVExcel(fileChooser.getSelectedFile().getName());
        
        if(ViewController.isPDF){
            SecurityManager secManager = new SecurityManager()
            {
                @Override
                public void checkExit(int status) {
                    throw new SecurityException();
                }
            };
            SecurityManager oldSecManager = System.getSecurityManager();
            System.setSecurityManager(secManager);
                    Thread t = new Thread(new Runnable() {
                       @Override
                       public void run() {
                            try {
                              ViewController.createCSV(fileChooser.getSelectedFile().getAbsolutePath());
                            } catch (SecurityException e) {
                                System.err.println("GOTCHA");
                            } catch (ParseException ex) {
                               Logger.getLogger(TableView.class.getName()).log(Level.SEVERE, null, ex);
                           }
                       }
                    });
                    t.start();
                    TableView tv = null;
                    try {
                    while(new File("C:\\Users\\Elpapo\\Desktop\\extractoPDF.cvs").canWrite())
                      tv = new TableView("C:\\Users\\Elpapo\\Desktop\\extractoPDF.cvs");
                      //tv.setTableModel(ViewController.fillVectorCSV("C:\\Users\\Elpapo\\Desktop\\Altman Exclusive Dealer Confidential Price List Jan  2014.pdf",1));
                      tv.setVisible(true);
                      this.dispose();
                    } catch (AccessControlException ace) { System.setSecurityManager(oldSecManager); System.err.println("oldSecManager");}
        }
        
        if(ViewController.isCSV){
            TableView tv = new TableView(fileChooser.getSelectedFile().getAbsolutePath());
            tv.setSheetName(fileChooser.getSelectedFile().getName());
                try {
                    tv.setTableModel(ViewController.populateTableCSV(fileChooser.getSelectedFile().getAbsolutePath()));
                } catch (FileNotFoundException ex) {
                    //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                }
             tv.setVisible(true);
             this.dispose();
        }
            
        if(ViewController.isExcel){
        switch(returnValue){
            case JFileChooser.APPROVE_OPTION:
                //File selectedFile = fileChooser.getSelectedFile();
                //System.out.println(fileChooser.getSelectedFile().getName());
                ViewController vc = null;
                try {
                    vc = new ViewController(fileChooser.getSelectedFile());
                } catch (IOException ioe) {
                    JOptionPane.showMessageDialog(this,
                                            "Ocurrio un error al obtener el Archivo.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
                    Toolkit.getDefaultToolkit().beep();
                    break;
                    //this.jButtonCargarActionPerformed(evt);
                    //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                } catch (OutOfMemoryError oome) {
                    JOptionPane.showMessageDialog(this,
                                            "El archivo es demasiado greande para leerlo,"
                                          + "\nno debe contener mas de 50,000 filas y 5000 columnas."
                                          + "\nSe recomienda abrir el archivo con el programa de origen"
                                          + "\npara dividirlo en extractos del archivo original.\n"
                                          + "Es posible que la memoria RAM este siendo utilizada"
                                          + "\npor programas robustos y no permitan el funcionamiento"
                                          + "\noptimo para la mineria de datos.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
                    Toolkit.getDefaultToolkit().beep();
                    break;
                }
                TableView tv = new TableView(fileChooser.getSelectedFile().getName());
                Object[] opc;
                String s = null;
                
                //System.err.println("num: " + vc.getNumSheet());
                 s = vc.getNameSheet(0);
                if(s != null){
                int index = 0;
                if( vc.getNumSheet() > 1){
                    opc = new Object[vc.getNumSheet()];
                    for(int i = 0; i < vc.getNumSheet();i++)
                        opc[i]=vc.getNameSheet(i);
                    s = (String)JOptionPane.showInputDialog(this,
                                            "Seleccione la Hoja de Calculo:\n",
                                            "Se encontraron varias Hojas.",
                                            JOptionPane.QUESTION_MESSAGE,
                                            null,opc,null);
                    if(s == null)
                        break;//s = vc.getNameSheet(0);
                    for(int i = 0; i < vc.getNumSheet(); i++)
                        if(s.equalsIgnoreCase(opc[i].toString()))
                            index = i;
                }
                tv.setSheetName(s);
                tv.setTableModel(vc.fillTableVector(index));
                tv.setVisible(true);
                this.dispose(); 
                }
                break;
            case JFileChooser.CANCEL_OPTION:
                //System.err.println("CancelOption");
                break;
            case JFileChooser.ERROR_OPTION:
                JOptionPane.showMessageDialog(this,
                                            "No fue posible obtener el archivo."
                                          + "\nVerifique que el archivo no este siendo usado por\n"
                                          + "otra persona/programa, este en vista protegida y/ó\n"
                                          + "cambiado de directorio/nombre.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
                //System.err.println("ErrorDesconocido");
                break;
            default:
                //System.err.println("Ocurrio un problema"+fileChooser.toString());
                break;
             } 
            }
        }
    }//GEN-LAST:event_jMenuOpenActionPerformed

    private void jButtonCabecerasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonCabecerasActionPerformed
        // TODO add your handling code here:
        int index = this.jTable1.getSelectedRow();
        Vector names = new Vector();
        if(this.jTable1.getRowSelectionAllowed() && this.isARowSelected()){
        for(int i = 0 ; i < this.jTable1.getColumnCount(); i++)
            //jTable1.getColumnModel().getColumn(i).setHeaderValue(this.jTable1.getValueAt(index, i).toString());
            names.add(this.jTable1.getValueAt(index, i).toString());
            ((CustomTableModel)this.jTable1.getModel()).setColumnIdentifiers(names);
            this.repaint();
//         if(jTable1.isRowSelected(i)){
//            jTable1.getColumnModel().getColumn(i).setHeaderValue(this.jTable1.getValueAt(index, i).toString());
//            this.repaint();
//         } else{ JOptionPane.showMessageDialog(this,
//                "Favor de seleccionar una fila\n Si la seleccion es multiple,\nse tomara la primera fila seleccionada.",
//                "No se encontraron filas seleccionadas.",
//                JOptionPane.ERROR_MESSAGE); break;}
            for(int i = 0;  i < this.jTable1.getColumnCount(); i++)
            this.jTable1.getColumnModel().getColumn(i).setMinWidth(100);
            this.repaint();
        } else { JOptionPane.showMessageDialog(this,
                "Esta opción funciona con la selección de al menos una fila\n"
                        + "\"Seleccion por Fila\" debe estar activado para realizar esta acción.",
                "No se encontro fila para selección.",
                JOptionPane.ERROR_MESSAGE);
                 Toolkit.getDefaultToolkit().beep();
        }
    }//GEN-LAST:event_jButtonCabecerasActionPerformed

    private void jButtonFilasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonFilasActionPerformed
        // TODO add your handling code here:
        if(this.bot1 == true){   
        this.jTable1.setColumnSelectionAllowed(true);
        this.jTable1.setRowSelectionAllowed(false);
        this.jButtonFilas.setText("<html><center>Seleccionando:<br>por Columna</center></html>");
        this.setTitle("Software Mineria- Seleccionando por Columna");
        this.bot1 = false;
        } else { 
            this.jTable1.setColumnSelectionAllowed(false);
            this.jTable1.setRowSelectionAllowed(true);
            this.jButtonFilas.setText("<html><center>Seleccionando:<br>por Fila</center></html>");
            this.bot1 = true;
        this.setTitle("Software Mineria- Seleccionando por Fila");    
        }
    }//GEN-LAST:event_jButtonFilasActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        int[] indices = null;
        if(this.jTable1.getRowSelectionAllowed()){
            indices = this.jTable1.getSelectedRows();
            Arrays.sort(indices);
            for (int i = indices.length - 1; i >= 0; i--){
                ((CustomTableModel)this.jTable1.getModel()).removeRow(indices[i]);
                ((CustomTableModel)this.jTable1.getModel()).fireTableRowsDeleted(indices[i],indices[i]);
            }
        } else {
            indices = this.jTable1.getSelectedColumns();
            Arrays.sort(indices);
            //for (int i = indices.length - 1; i >= 0; i--) 
            //((CustomTableModel)this.jTable1.getModel()).setColumnCount(this.jTable1.getColumnCount()-1);
            //this.jTable1.setModel(ViewController.deleteColumn(this.jTable1,indices));
            ViewController.deleteColumn(jTable1, indices);
            //this.repaint();
        }
        //fireTableRowsDeleted(indices[i], indices[i]);
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton7ActionPerformed
        /*try {
            // TODO add your handling code here:
            if(this.colIndex.length > 1){
             ViewController.saveData(this.colIndex,this.jTable1);
             JOptionPane.showMessageDialog(this,
                "Datos guardados satisfactoriamente",
                "Guardado completo",
                JOptionPane.INFORMATION_MESSAGE);
            } else JOptionPane.showMessageDialog(this,
                "No se han seleccionado los indices de las columnas.",
                "Error en la Seleccion de Columnas.",
                JOptionPane.ERROR_MESSAGE);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                                            "No fue posible guardar el archivo.\n"
                                          + "Verifique si el archivo no este siendo\n"
                                          + "usado por otra persona u otro programa.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
            //Logger.getLogger(TableView.class.getName()).log(Level.SEVERE, null, ex);
        }*/
        if(jComboBox1.getSelectedIndex()!=0)
        try {   
        
        if(this.colIndex != null || this.colIndex.length > 0 ){
        /*Thread t = new Thread(new Runnable() {
                    @Override
                    public void run() {
                      sw.setVisible(true);
                    }
         });
         t.start();*/
          String sFileName = "", sFilePath = "";
          final JFileChooser fileChooser = new JFileChooser();
          fileChooser.setFileFilter(new FileNameExtensionFilter("MS-Excel", "xlsx","xls"));
          fileChooser.setDialogTitle("Seleccione a que documento adjuntar los datos.");
          fileChooser.setApproveButtonText("Adjuntar");
          
          int jPanelOption;
          int status = fileChooser.showSaveDialog(this);
          if(JFileChooser.APPROVE_OPTION == status){
            sFileName = fileChooser.getSelectedFile().getName();
            sFilePath = fileChooser.getCurrentDirectory().getPath();
            if(sFileName.toLowerCase().endsWith("xlsx")){
             this.openOptionPanel();
             if(this.factores != null){
                 ViewController.factores = this.factores;
                 this.factores = null;
             }
             if(fileChooser.getSelectedFile().exists()){
                Toolkit.getDefaultToolkit().beep();
                jPanelOption = JOptionPane.showOptionDialog(null,
                            "¡Se agregaran los datos al final del documento!.\n"
                                    + "¿Desea continuar con esta acción?.",
                            "Acción peligrosa!",
                            JOptionPane.YES_NO_OPTION
                            ,JOptionPane.WARNING_MESSAGE,
                            null,null,null);
                
                if(jPanelOption == JOptionPane.OK_OPTION){//Llamar escribir archivo
                     this.saveDataSwingWorker(sFileName, sFilePath, false);
                    //System.err.println(sFileName+" -- "+sFilePath+"  "+jPanelOption);
                }
             } else {
                Toolkit.getDefaultToolkit().beep();
                JOptionPane.showMessageDialog(this,
                "Esta opcion solo puede adjuntar datos a\narchivos Excel existentes.\n"
              + "Si desea GUARDAR un archivo nuevo, use la opcion en la barra de menu.",
                "No existe ese archivo.",
                JOptionPane.ERROR_MESSAGE);
             } 
            } else {
                Toolkit.getDefaultToolkit().beep();
                JOptionPane.showMessageDialog(this,
                "Solo es posible adjuntar datos a\narchivos Excel versión 2007 y posteriores.",
                "Incompatibilidad de documento",
                JOptionPane.ERROR_MESSAGE);
            }
          } 
                
         //sw.saveData(this.colIndex,this.jTable1);
         
         //sw.setVisible(false);
         //sw.dispose();
        } /*else { JOptionPane.showMessageDialog(this,
                "No se han seleccionado los indices de las columnas.",
                "Error en la Seleccion de Columnas.",
                JOptionPane.ERROR_MESSAGE);
                Toolkit.getDefaultToolkit().beep();
           }*/
        
        } catch (NullPointerException npe) {
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                "No se han seleccionado los indices de las columnas.",
                "Error en la Seleccion de Columnas.",
                JOptionPane.ERROR_MESSAGE);
        } else{ 
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                "No se ha seleccionado una marca, favor de seleccionar una"
                        + "\nopcion del menu \"Marcas\"",
                "Error en la Seleccion de Marcas.",
                JOptionPane.ERROR_MESSAGE);
                
        }
  
    }//GEN-LAST:event_jButton7ActionPerformed

    private void openOptionPanel(){
        /*NumberFormat format =  DecimalFormat.getInstance();
        format.setMinimumFractionDigits(0);
        format.setMaximumFractionDigits(3);
        format.setMinimumIntegerDigits(0);
        format.setMaximumIntegerDigits(3);
        NumberFormatter formatter = new NumberFormatter(format);
        //formatter.setValueClass(Float.class);
        //formatter.setMinimum(0.0);
        //formatter.setMaximum(Float.MAX_VALUE);
        // If you want the value to be committed on each keystroke instead of focus lost
        //formatter.setCommitsOnValidEdit(true);
        formatter.setAllowsInvalid(false);*/
        JFormattedTextField field1 = new JFormattedTextField();//formatter);
        JFormattedTextField field2 = new JFormattedTextField();//formatter);
        JFormattedTextField field3 = new JFormattedTextField();//formatter);
        JFormattedTextField field4 = new JFormattedTextField();//formatter);

        Object[] message = {
            "Factor Publico:\n", field1,
            "Factor Contratista:\n", field2,
            "Factor Mayoreo:\n", field3,
            "Factor Super Mayoreo:\n", field4
        };
        
       int option = JOptionPane.showConfirmDialog(this, message, "Factores para la marca: "+this.jComboBox1.getSelectedItem(), JOptionPane.OK_CANCEL_OPTION);
    
       if (option == JOptionPane.OK_OPTION){
        try {
        this.factores = new float[4];
        factores[0] = Float.parseFloat(field1.getText().trim().replace("", "1"));
        factores[1] = Float.parseFloat(field2.getText().trim().replace("", "1"));
        factores[2] = Float.parseFloat(field3.getText().trim().replace("", "1"));
        factores[3] = Float.parseFloat(field4.getText().trim().replace("", "1"));
        } catch (NumberFormatException e){
           this.factores = null; 
           Toolkit.getDefaultToolkit().beep();
           JOptionPane.showMessageDialog(this,
                "Favor de introducir solo numeros",
                "Error de entrada",
                JOptionPane.INFORMATION_MESSAGE);
           openOptionPanel();
        }
       }   
   }
    
    public void saveData( int[] colIndex, JTable jTable1 , JComboBox marca, boolean se) {
        try {
            // TODO add your handling code here:
             ViewController.saveData(colIndex,jTable1,marca,se);
             JOptionPane.showMessageDialog(this,
                "Datos guardados satisfactoriamente",
                "Guardado completo",
                JOptionPane.INFORMATION_MESSAGE);
             //this.setVisible(false);
             //this.dispose();
        } catch (Exception ex) {
            ex.printStackTrace();
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                                            "No fue posible guardar el archivo.\n"
                                          + "Verifique si el archivo no este siendo\n"
                                          + "usado por otra persona u otro programa.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
            //Logger.getLogger(TableView.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public void saveData( int[] colIndex, JTable jTable1 , JComboBox marca, String sFileName, String sFilePath) {
        try {
            // TODO add your handling code here:
             ViewController.saveData(colIndex,jTable1,marca,sFileName,sFilePath);
             Toolkit.getDefaultToolkit().beep();
             JOptionPane.showMessageDialog(this,
                "Datos guardados satisfactoriamente",
                "Guardado completo",
                JOptionPane.INFORMATION_MESSAGE);
             //this.setVisible(false);
             //this.dispose();
        } catch (Exception ex) {
            ex.printStackTrace();
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                                            "No fue posible guardar el archivo.\n"
                                          + "Verifique si el archivo no este siendo\n"
                                          + "usado por otra persona u otro programa.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
            
            //Logger.getLogger(TableView.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private void saveDataSwingWorker(final String sFileName, final String sFilePath, final boolean sobreExcribir){
        final SaveWindow sw = new SaveWindow(this, true);
         SwingWorker worker;
               worker = new SwingWorker<Void, Void>() {
                                   
                   @Override
                   public Void doInBackground() {
                        try {
                            //============> Aqui se guardan los datos
                            saveData(colIndex, jTable1, jComboBox1, sFileName, sFilePath);
                            sw.getPB().setIndeterminate(false);
                            sw.getPB().setValue(100);
                            Thread.sleep(300);
                            
                            //this.setVisible(false);
                            //this.dispose();
                        } catch (Exception ex) {
                            ex.printStackTrace();
                            Toolkit.getDefaultToolkit().beep();
                            JOptionPane.showMessageDialog(null,
                                            "No fue posible guardar el archivo.\n"
                                          + "Verifique si el archivo no este siendo\n"
                                          + "usado por otra persona u otro programa.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
                            
                            //Logger.getLogger(TableView.class.getName()).log(Level.SEVERE, null, ex);
                            }
                       return null;
                   }
                   
                   @Override
                   public void done() {
                       //Toolkit.getDefaultToolkit().beep();
                       sw.dispose();
                       
                   }};
               worker.execute();
         sw.setVisible(true);
         
         if(worker.isDone()){
           sw.dispose();
           //System.err.println("DONE");
         }
    }
    
    private void jButtonSeleccionColumnActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonSeleccionColumnActionPerformed
        // TODO add your handling code here:
        
        final MarkColumns cc = new MarkColumns(this, true);
        cc.setModalityType(Dialog.ModalityType.DOCUMENT_MODAL);
        //cc.setVisible(true);
        
        CustomTableModel ctm = new CustomTableModel();
        Vector header = new Vector();
        Vector data = new Vector();
        Vector row;
        
        for(int i = 0; i < this.jTable1.getModel().getColumnCount(); i++){
         header.add(this.jTable1.getModel().getColumnName(i));
         this.jTable1.getColumnModel().getColumn(i).setMinWidth(100);
        }
        if(this.jTable1.getModel().getRowCount() > 10)
        for(int i = 0; i < 6; i++){
            row = new Vector();
            for(int j = 0; j < this.jTable1.getModel().getColumnCount(); j++){
             row.add(this.jTable1.getModel().getValueAt(i, j));          
             //System.out.println(this.jTable1.getModel().getValueAt(i, j));
            }
            data.add(row);
        }
        
        else 
            for(int i = 0; i < this.jTable1.getModel().getRowCount() ; i++){
            row = new Vector();
            for(int j = 0; j < this.jTable1.getModel().getColumnCount(); j++){
             row.add(this.jTable1.getModel().getValueAt(i, j));          
             //System.out.println(this.jTable1.getModel().getValueAt(i, j));
            }
            data.add(row);
        }
            
        
        ctm.setDataVector(data, header);
        cc.setTableModel(ctm);
        cc.setVisible(true);
        cc.repaint();
        colIndex = cc.getIndexes(); //0 no seleccionado
        cc.dispose();
        if(colIndex == null)
            colIndex = null;//new int[0];
        /*for(int i = 0; i < colIndex.length; i ++){
            //System.out.println(((JComboBox)this.jPanelComboBox.getComponent(i)).getSelectedIndex());
            System.out.print(colIndex[i]+"-");
         }*/
        
        
        /*Object[] possibilities = {"ham", "spam", "yam"};
        String s = (String)JOptionPane.showInputDialog(
                    this,
                    "Seleccione la Hoja de Calculo:\n",
                    "Se encontraron varias Hojas.",
                    JOptionPane.QUESTION_MESSAGE,
                    null,
                    possibilities,
                    null);
         */
        //System.out.println(s);
    }//GEN-LAST:event_jButtonSeleccionColumnActionPerformed

    private void jMenuItem1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem1ActionPerformed
        // TODO add your handling code here:
        JOptionPane.showMessageDialog(this,
                "/*\n" +
"  Copyright 2015 Isaac Alcocer <aosi87@gmail.com>.\n" +
" \n" +
"  Licensed under the Apache License, Version 2.0 (the \"License\");\n" +
"  you may not use this file except in compliance with the License.\n" +
"  You may obtain a copy of the License at\n" +
" \n" +
"       http://www.apache.org/licenses/LICENSE-2.0\n" +
" \n" +
"  Unless required by applicable law or agreed to in writing, software\n" +
"  distributed under the License is distributed on an \"AS IS\" BASIS,\n" +
"  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\n" +
"  See the License for the specific language governing permissions and\n" +
"  limitations under the License.\n" +
" ",
                "Version del programa: RC 1.0.7504",
                JOptionPane.INFORMATION_MESSAGE);
        
    }//GEN-LAST:event_jMenuItem1ActionPerformed

    private void jMenuItem3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItem3ActionPerformed
        // TODO add your handling code here:
        try {
        //checamos si ya se selecciono la marca/giro
        if(this.jComboBox1.getSelectedIndex() != 0){
        //checamos si ya hay columnas seleccionadas
        if(this.colIndex.length > 0 || this.colIndex != null){
          String sFileName = "", sFilePath = "";
          final JFileChooser fileChooser = new JFileChooser();
          fileChooser.setFileFilter(new FileNameExtensionFilter("MS-Excel", "xlsx","xls"));
          int jPanelOption;
          int status = fileChooser.showSaveDialog(this);
          if(JFileChooser.APPROVE_OPTION == status){
            sFileName = fileChooser.getSelectedFile().getName();
            sFilePath = fileChooser.getCurrentDirectory().getPath();
            if(fileChooser.getSelectedFile().exists()){
                Toolkit.getDefaultToolkit().beep();
                jPanelOption = JOptionPane.showOptionDialog(null,
                            "¡El nombre del archivo ya existe!.\n"
                                    + "¿Desea sobrescribirlo?.",
                            "El nombre del archivo ya EXISTE",
                            JOptionPane.YES_NO_OPTION
                            ,JOptionPane.WARNING_MESSAGE,
                            null,null,null);
                
                if(jPanelOption == 0){//Llamar escribir archivo
                     this.saveDataSwingWorker(sFileName, sFilePath, true);
                    //System.err.println(sFileName+" -- "+sFilePath+"  "+jPanelOption);
                }
            } else {//adjuntar
                     this.saveDataSwingWorker(sFileName, sFilePath, true);
                     //System.out.println("crear");
                   } 
          }
        
        } else {
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                "No se ha seleccionado una marca, favor de seleccionar una"
                        + "\nopcion del menu \"Marcas\"",
                "Error en la Seleccion de Marcas.",
                JOptionPane.ERROR_MESSAGE);
                
           }//fin if columnas seleccionadas
        
        } else {
            Toolkit.getDefaultToolkit().beep();
                JOptionPane.showMessageDialog(this,
                "No se ha seleccionado una marca, favor de seleccionar una"
                        + "\nopcion del menu \"Marcas\"",
                "Error en la Seleccion de Marcas.",
                JOptionPane.ERROR_MESSAGE);
                
                }//fin if marca
        
        } catch (NullPointerException npe) {
            Toolkit.getDefaultToolkit().beep();
            JOptionPane.showMessageDialog(this,
                "No se han seleccionado los indices de las columnas.",
                "Error en la Seleccion de Columnas.",
                JOptionPane.ERROR_MESSAGE);
            
          }
        System.out.println("si - guardar como");
    }//GEN-LAST:event_jMenuItem3ActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton7;
    private javax.swing.JButton jButtonCabeceras;
    private javax.swing.JButton jButtonFilas;
    private javax.swing.JButton jButtonSeleccionColumn;
    private javax.swing.JCheckBox jCheckBox1;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItem3;
    private javax.swing.JMenuItem jMenuItemExit;
    private javax.swing.JMenuItem jMenuMarcas;
    private javax.swing.JMenuItem jMenuOpen;
    private javax.swing.JPanel jPanelCombo;
    private javax.swing.JPanel jPanelManejoDatos;
    private javax.swing.JPanel jPanelManejoTabla;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JPopupMenu.Separator jSeparator1;
    private javax.swing.JPopupMenu.Separator jSeparator2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
