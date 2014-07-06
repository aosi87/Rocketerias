/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package view;

import controller.Controller;
import javax.swing.BorderFactory;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.border.TitledBorder;
import javax.swing.table.DefaultTableModel;

/**
 *
 * @author Isaac Alcocer <aosi87@gmail.com>
 */
public class TableView extends javax.swing.JFrame {
    
    private boolean bot1 = false;
    /**
     * Creates new form TableView
     */
    public TableView() {
        initComponents();
    }
    
    public TableView(String path) {
        initComponents();
        this.jTabbedPane1.setTitleAt(0, path);
        this.setLocationRelativeTo(null);
        this.jPanelManejoTabla.setBorder(null);
        this.jTable1.setCellSelectionEnabled(true);
        this.jPanelManejoDatos.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            "Manejo de Datos",
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
        this.jPanelManejoTabla.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder (),
                                                            "Manejo de Tabla",
                                                            TitledBorder.CENTER,
                                                            TitledBorder.TOP));
        //this.jTable1.setColumnSelectionAllowed(false);
        //this.jTable1.setRowSelectionAllowed(false);
        
        
    }
    
    public void setTableModel(DefaultTableModel model){
        this.jTable1.setModel(model);
        this.jTable1.setValueAt(true, 0, 0);
        this.repaint();
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
        jButton7 = new javax.swing.JButton();
        jPanelManejoDatos = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenuOpen = new javax.swing.JMenuItem();
        jSeparator1 = new javax.swing.JPopupMenu.Separator();
        jMenuItemExit = new javax.swing.JMenuItem();
        jMenu2 = new javax.swing.JMenu();
        jMenuItem1 = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setMaximumSize(new java.awt.Dimension(1080, 720));
        setMinimumSize(new java.awt.Dimension(800, 600));
        setName("TableView"); // NOI18N

        jTabbedPane1.setBorder(javax.swing.BorderFactory.createBevelBorder(javax.swing.border.BevelBorder.RAISED));
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
        jScrollPane1.setViewportView(jTable1);

        jTabbedPane1.addTab("Archivo", jScrollPane1);

        jPanelManejoTabla.setLayout(new java.awt.GridLayout());

        jButtonFilas.setText("<html><center>Seleccionar<br>Filas/Columnas</center></html>");
        jButtonFilas.setMaximumSize(new java.awt.Dimension(87, 37));
        jButtonFilas.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonFilasActionPerformed(evt);
            }
        });
        jPanelManejoTabla.add(jButtonFilas);

        jButtonCabeceras.setText("<html>Seleccionar <br>Cabeceras</html>");
        jButtonCabeceras.setMaximumSize(new java.awt.Dimension(87, 37));
        jButtonCabeceras.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonCabecerasActionPerformed(evt);
            }
        });
        jPanelManejoTabla.add(jButtonCabeceras);

        jButton7.setText("<html><center>Adjuntar<br>Datos</center></html>");
        jPanelManejoTabla.add(jButton7);

        jPanelManejoDatos.setLayout(new java.awt.GridLayout());

        jButton1.setText("<html><center>Borrar<br>Selección</center></html>");
        jPanelManejoDatos.add(jButton1);

        jButton2.setText("<html><center>Mantener<br>Seleccion</center></html>");
        jPanelManejoDatos.add(jButton2);

        jMenu1.setText("Archivo");

        jMenuOpen.setText("Abrir");
        jMenuOpen.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jMenuOpenActionPerformed(evt);
            }
        });
        jMenu1.add(jMenuOpen);
        jMenu1.add(jSeparator1);

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
        jMenu2.add(jMenuItem1);

        jMenuBar1.add(jMenu2);

        setJMenuBar(jMenuBar1);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jTabbedPane1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanelManejoTabla, javax.swing.GroupLayout.PREFERRED_SIZE, 309, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 190, Short.MAX_VALUE)
                        .addComponent(jPanelManejoDatos, javax.swing.GroupLayout.PREFERRED_SIZE, 194, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jPanelManejoTabla, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jPanelManejoDatos, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
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
        JFileChooser fileChooser = new JFileChooser();
        //Agregamos un filtro de extensiones
        //fileChooser.setFileFilter(new FileNameExtensionFilter("Doc - MS-Office 2003", "doc"));
        int returnValue = fileChooser.showDialog(null,"Seleccionar");
        switch(returnValue){
            case JFileChooser.APPROVE_OPTION:
                this.setTableModel(new Controller(fileChooser.getSelectedFile()).fillTableVector());
                break;
            case JFileChooser.CANCEL_OPTION:
                System.err.println("CancelOption");
                break;
            case JFileChooser.ERROR_OPTION:
                System.err.println("ErrorDesconocido");
                break;
            default:
                System.err.println("Ocurrio un problema"+fileChooser.toString());
                break;
        }
    }//GEN-LAST:event_jMenuOpenActionPerformed

    private void jButtonCabecerasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonCabecerasActionPerformed
        // TODO add your handling code here:
        int index = this.jTable1.getSelectedRow();
        if(!this.jTable1.getColumnSelectionAllowed()){
        for(int i = 0 ; i < this.jTable1.getColumnCount(); i++)
         jTable1.getColumnModel().getColumn(i).setHeaderValue(this.jTable1.getValueAt(index, i).toString());
        this.repaint();
        }
        else JOptionPane.showMessageDialog(this,
                "Esta opción funciona con seleccion de fila.\nFavor de cambiar el tipo de selección.",
                "Se encontro una columna!!!",
                JOptionPane.ERROR_MESSAGE);
    }//GEN-LAST:event_jButtonCabecerasActionPerformed

    private void jButtonFilasActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonFilasActionPerformed
        // TODO add your handling code here:
        if(this.bot1 == false){
        //this.jTable1.setCellSelectionEnabled(false);    
        this.jTable1.setColumnSelectionAllowed(false);
        this.jTable1.setRowSelectionAllowed(true);
        this.jButtonFilas.setText("<html><center>Selección<br>por Fila</center></html>");
        this.bot1 = true;
        } else {
            //this.jTable1.setCellSelectionEnabled(true);    
            this.jTable1.setColumnSelectionAllowed(true);
            this.jTable1.setRowSelectionAllowed(false);
            this.jButtonFilas.setText("<html><center>Selección<br>por Columna</center></html>");
            this.bot1 = false;
        }
    }//GEN-LAST:event_jButtonFilasActionPerformed

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton7;
    private javax.swing.JButton jButtonCabeceras;
    private javax.swing.JButton jButtonFilas;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuItem jMenuItem1;
    private javax.swing.JMenuItem jMenuItemExit;
    private javax.swing.JMenuItem jMenuOpen;
    private javax.swing.JPanel jPanelManejoDatos;
    private javax.swing.JPanel jPanelManejoTabla;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JPopupMenu.Separator jSeparator1;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
}
