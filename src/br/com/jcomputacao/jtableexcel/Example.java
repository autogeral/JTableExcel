/*
 * Example.java
 *
 * Created on 16/09/2011, 11:43:56
 */
package br.com.jcomputacao.jtableexcel;

import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.Window;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;

/**
 *
 * @author Murilo
 */
public class Example extends javax.swing.JFrame {

    /** Creates new form Example */
    public Example() {
        initComponents();
        configureColumns();
    }

    private void configureColumns() {
        int i = 0;
        TableColumn col = table.getColumnModel().getColumn(i++);
        col.setPreferredWidth(240);
        col = table.getColumnModel().getColumn(i++);
        col.setPreferredWidth(80);
        col = table.getColumnModel().getColumn(i++);
        col.setPreferredWidth(120);
        col = table.getColumnModel().getColumn(i++);
        col.setPreferredWidth(80);
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        scrollPane = new javax.swing.JScrollPane();
        table = new javax.swing.JTable();
        bExport = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("IT Mega Companies in the World - This is a JOKE!!");

        table.setModel(createTableModel());
        table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        scrollPane.setViewportView(table);

        getContentPane().add(scrollPane, java.awt.BorderLayout.CENTER);

        bExport.setText("Export");
        bExport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                bExportActionPerformed(evt);
            }
        });
        getContentPane().add(bExport, java.awt.BorderLayout.PAGE_START);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void bExportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bExportActionPerformed
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileFilter() {

            @Override
            public boolean accept(File f) {
                //return (f.getName().toLowerCase().endsWith(".xlsx") || f.getName().toLowerCase().endsWith(".xls"));
                return (f.getName().toLowerCase().endsWith(".xls"));
            }

            @Override
            public String getDescription() {
                return "Excel file";
            }
        });

        int status = fc.showSaveDialog(this);
        if (status == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            OutputStream os = null;
            try {
                os = new FileOutputStream(file);
                ExcelExporter ee = new ExcelExporter(this.table.getModel(), os);
                ee.execute();
            } catch (IOException ex) {
                Logger.getLogger(Example.class.getName()).log(Level.SEVERE, null, ex);
                JOptionPane.showMessageDialog(this, 
                        "Error when trying to write the table to excel file", 
                        "Error", JOptionPane.ERROR_MESSAGE);
            } finally {
                if (os != null) {
                    try {
                        os.close();
                    } catch (IOException ex) {
                        Logger.getLogger(Example.class.getName()).log(Level.SEVERE, "Error writting to file", ex);
                    }
                }
            }
            JOptionPane.showMessageDialog(this, 
                        "Excel worksheet created on "+file.getAbsoluteFile().getName(), 
                        "Worksheet created", JOptionPane.INFORMATION_MESSAGE);
        }

    }//GEN-LAST:event_bExportActionPerformed

    private javax.swing.table.DefaultTableModel createTableModel() {
        Object[][] data = new Object[][]{
            {"J Computa\u00e7\u00e3o", "Brazil", 777.2d, 2},
            {"Oracle", "USA", 541.9, 31242},
            {"IBM", "USA", 437.1, 19241},
            {"Microsoft", "USA", 341.5, 18427},};
        String[] colNames = new String[]{"Name", "HQ Country", "Anual Revenue (Billions)", "Employees"};

        return new AnotherTableModel(data, colNames);
    }

    private class AnotherTableModel extends DefaultTableModel {

        private AnotherTableModel(Object[][] data, String[] colNames) {
            super(data, colNames);
        }

        @Override
        public Class<?> getColumnClass(int columnIndex) {
            if (columnIndex == 2) {
                return Double.class;
            } else if (columnIndex == 3) {
                return Integer.class;
            }
            return String.class;
        }
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(new Runnable() {

            @Override
            public void run() {
                Example example = new Example();
                example.setSize(600, 300);
                center(example);
                example.setVisible(true);
            }
        });
    }

    private static void center(Window window) {
        Dimension myDim = window.getSize();
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        if (myDim.height > (dim.height - 30)) {
            myDim.height = dim.height - 30;
            window.setSize(myDim);
        }

        if (myDim.width > dim.width) {
            myDim.width = dim.width;
            window.setSize(myDim);
        }

        int x = (dim.width / 2 - myDim.width / 2);
        int y = (dim.height / 2 - myDim.height / 2);
        window.setLocation(x, y);
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton bExport;
    private javax.swing.JScrollPane scrollPane;
    private javax.swing.JTable table;
    // End of variables declaration//GEN-END:variables
}
