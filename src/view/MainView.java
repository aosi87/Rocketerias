/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package view;

//import com.alee.laf.WebLookAndFeel;
import controller.Tabula;
import controller.ViewController;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.commons.cli.ParseException;

/**
 *
 * @author Isaac Alcocer
 */
public class MainView extends javax.swing.JFrame {

    /**
     * Creates new form MainView
     */
    public MainView() {
        initComponents();
        this.setLocationRelativeTo(null);
        this.setResizable(false);
        this.jLabelLogo.setText("");
        try {
            this.jLabelLogo.setIcon(new ImageIcon(this.loadImageIntoJLABEL()));
        } catch (IOException ex) {
            //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private BufferedImage loadImageIntoJLABEL() throws IOException{
            BufferedImage bi=new BufferedImage(jLabelLogo.getWidth(),jLabelLogo.getHeight(),BufferedImage.TYPE_INT_ARGB);
            Graphics2D g=bi.createGraphics();
            
            Image img = ImageIO.read(getClass().getClassLoader().getResource("gfx/logoMining.png"));//new File("resources/gfx/logoMining.png"));
            g.drawImage(img, 0, 0, jLabelLogo.getWidth(), jLabelLogo.getHeight(), null);
            g.dispose();
            return bi;
    }
    
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jLabelLogo = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();
        jButtonCargar = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jButtonExit = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Cargar lista de productos.");
        setIconImage(new ImageIcon(getClass().getClassLoader().getResource("gfx/icon.png")).getImage());

        jPanel1.setLayout(new java.awt.BorderLayout());

        jLabelLogo.setFont(new java.awt.Font("Tahoma", 0, 36)); // NOI18N
        jLabelLogo.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabelLogo.setText("Logo");
        jPanel1.add(jLabelLogo, java.awt.BorderLayout.CENTER);

        jPanel2.setLayout(new java.awt.GridLayout(1, 0));

        jButtonCargar.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jButtonCargar.setText("Cargar archivo");
        jButtonCargar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonCargarActionPerformed(evt);
            }
        });
        jPanel2.add(jButtonCargar);

        jPanel3.setLayout(new java.awt.GridLayout(1, 0));

        jButtonExit.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jButtonExit.setText("Salir");
        jButtonExit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButtonExitActionPerformed(evt);
            }
        });
        jPanel3.add(jButtonExit);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, 159, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, 100, Short.MAX_VALUE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, 72, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, 64, Short.MAX_VALUE)
                    .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButtonCargarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonCargarActionPerformed
        JFileChooser fileChooser = new JFileChooser();
        //Agregamos un filtro de extensiones
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("MS-Word", "docx","doc"));
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Adobe-PDF", "pdf"));
        fileChooser.setFileFilter(new FileNameExtensionFilter("MS-Excel", "xlsx","xls"));
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        
        int returnValue = fileChooser.showDialog(this,"Seleccionar documento");
        if(fileChooser.getSelectedFile() != null){
            ViewController.setPDFCSVExcel(fileChooser.getSelectedFile().getName());
        
        if(ViewController.isPDF){
                try {
                    ViewController.createCSV(fileChooser.getSelectedFile().getAbsolutePath());
                    new TableView();
                    this.dispose();
                } catch (ParseException ex) {
                    Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                }
        }
        
        if(ViewController.isCSV){
            TableView tv = new TableView(fileChooser.getSelectedFile().getAbsolutePath());
            tv.setSheetName(fileChooser.getSelectedFile().getName());
                try {
                    tv.setTableModel(ViewController.fillTableCSV(ViewController.insertDataCSV(fileChooser.getSelectedFile().getAbsolutePath())));
                    
                } catch (FileNotFoundException ex) {
                    //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IOException ex) {
                    //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                }
             tv.setVisible(true);
             tv.repaint();
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
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(this,
                                            "Ocurrio un error al obtener el Archivo.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
                    break;
                    //this.jButtonCargarActionPerformed(evt);
                    //Logger.getLogger(MainView.class.getName()).log(Level.SEVERE, null, ex);
                } catch (OutOfMemoryError oome) {
                    JOptionPane.showMessageDialog(this,
                                            "El archivo es demasiado greande para leerlo"
                                          + "\ncon menos de 2Gigs de memoria utilizable."
                                          + "\nSe recomienda dividir el archivo en archivos\n"
                                          + "mas pequeños para su manejo optimo.",
                                            "ERROR",
                                            JOptionPane.ERROR_MESSAGE);
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
    }//GEN-LAST:event_jButtonCargarActionPerformed

    private void jButtonExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButtonExitActionPerformed
        // TODO add your handling code here:
        //simplemente cerramos la instancia, o la aplicacion por completo.
        this.dispose();
    }//GEN-LAST:event_jButtonExitActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
               
        /*
        try { 
         // Setting up WebLookAndFeel style
            UIManager.setLookAndFeel ( WebLookAndFeel.class.getCanonicalName () );
        }
            catch ( Throwable e )
        {
            // Something went wrong
        }*/
        
        // Set the Nimbus look and feel 
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    //javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            //java.util.logging.Logger.getLogger(MainView.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            //java.util.logging.Logger.getLogger(MainView.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            //java.util.logging.Logger.getLogger(MainView.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            //java.util.logging.Logger.getLogger(MainView.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        
        
        // Set WebLAF look and feel 
         //WebLookAndFeel.install();
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                new MainView().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButtonCargar;
    private javax.swing.JButton jButtonExit;
    private javax.swing.JLabel jLabelLogo;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    // End of variables declaration//GEN-END:variables

}