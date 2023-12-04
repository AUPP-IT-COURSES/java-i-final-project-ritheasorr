package com.kingaspx.main;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.github.sarxos.webcam.Webcam;
import com.github.sarxos.webcam.WebcamPanel;
import com.github.sarxos.webcam.WebcamResolution;
import com.google.zxing.BinaryBitmap;
import com.google.zxing.LuminanceSource;
import com.google.zxing.MultiFormatReader;
import com.google.zxing.NotFoundException;
import com.google.zxing.Result;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.common.HybridBinarizer;
import java.awt.Dimension;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.util.concurrent.Executor;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;
import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Menu extends javax.swing.JFrame implements Runnable, ThreadFactory {

    private WebcamPanel panel = null;
    private Webcam webcam = null;

    private static final long serialVersionUID = 6441489157408381878L;
    private Executor executor = Executors.newSingleThreadExecutor(this);

    //new
    private Workbook workbook;
    private Sheet sheet;
    private int rowIndex;
    private String filePath = "output2.xlsx";

    public Menu() {
        initComponents();
        initWebcam();
        initializeExcel();


        try {
            workbook = new XSSFWorkbook(new FileInputStream(filePath));
            sheet = workbook.getSheet("QR Code Results");
            if (sheet == null) {
                sheet = workbook.createSheet("QR Code Results");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }




    //new
    private void initializeExcel() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("QR Code Results");
        rowIndex = 0;

        // Create header row
        Row headerRow = sheet.createRow(rowIndex++);
        headerRow.createCell(0).setCellValue("QR Code Result");

        // Set up column widths (optional)
        sheet.setColumnWidth(0, 10000);
    }
//    private void writeResultToExcel(String result) {
//        Row row = sheet.createRow(rowIndex++);
//        row.createCell(0).setCellValue(result);}
    private void writeResultToExcel(String result) {
        int lastRowNum = sheet.getLastRowNum();
        Row row = sheet.createRow(lastRowNum + 1);
        Cell cell = row.createCell(0);
        cell.setCellValue(result);

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            System.out.println("QR code result appended to Excel file: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @SuppressWarnings("unchecked")

    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        result_field = new javax.swing.JTextField();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel1 = new javax.swing.JLabel();
        jPanel2 = new javax.swing.JPanel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

//
//        saveButton = new javax.swing.JButton();
//        saveButton.setText("Save to Excel");
//        saveButton.addActionListener(new java.awt.event.ActionListener() {
//            public void actionPerformed(java.awt.event.ActionEvent evt) {
//                saveButtonActionPerformed(evt);
//            }
//        });
//        jPanel1.add(saveButton, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 380, 120, -1));


        jPanel1.setBackground(new java.awt.Color(255, 255, 255));
        jPanel1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        result_field.setBorder(null);
        jPanel1.add(result_field, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 340, 470, 20));

        jSeparator1.setForeground(new java.awt.Color(126, 167, 206));
        jPanel1.add(jSeparator1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 360, 470, 10));

        jLabel1.setForeground(new java.awt.Color(105, 105, 105));
        jLabel1.setText("Result");
        jPanel1.add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 320, -1, -1));

        jPanel2.setBackground(new java.awt.Color(250, 250, 250));
        jPanel2.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(230, 230, 230)));
        jPanel2.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel1.add(jPanel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 10, 470, 300));

        getContentPane().add(jPanel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 500, 380));

        pack();
    }// </editor-fold>//GEN-END:initComponents

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Menu.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        //</editor-fold>

        /* Create and display the form */
        SwingUtilities.invokeLater(() -> {
            new Menu().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel jLabel1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JTextField result_field;
    // End of variables declaration//GEN-END:variables
    private javax.swing.JButton saveButton;



    private void initWebcam() {
        Dimension size = WebcamResolution.QVGA.getSize();
        webcam = Webcam.getWebcams().get(0); //0 is default webcam
        webcam.setViewSize(size);

        panel = new WebcamPanel(webcam);
        panel.setPreferredSize(size);
        panel.setFPSDisplayed(true);

        jPanel2.add(panel, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 470, 300));

        executor.execute(this);
    }

    @Override
    public void run() {
        do {
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            Result result = null;
            BufferedImage image = null;

            if (webcam.isOpen()) {
                if ((image = webcam.getImage()) == null) {
                    continue;
                }
            }

            LuminanceSource source = new BufferedImageLuminanceSource(image);
            BinaryBitmap bitmap = new BinaryBitmap(new HybridBinarizer(source));

            try {
                result = new MultiFormatReader().decode(bitmap);
            } catch (NotFoundException e) {
                //No result...
            }
            if (result != null) {
                result_field.setText(String.valueOf(result));
                writeResultToExcel(result.getText());
//                writeResultToExcel(String.valueOf(result)); // Write result to Excel
            }
//            if (result != null) {
//                result_field.setText(result.getText());
//            }
        } while (true);
    }

    public void stop() {
        // Other code...

        try {
            FileOutputStream fileOutputStream = new FileOutputStream(filePath);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public Thread newThread(Runnable r) {
        Thread t = new Thread(r, "My Thread");
        t.setDaemon(true);
        return t;
    }




//    private void saveButtonActionPerformed(java.awt.event.ActionEvent evt) {
//        saveExcelFile();
//    }
    //new
//    private void saveExcelFile() {
//        try {
//            // Choose a file to save
//            JFileChooser fileChooser = new JFileChooser();
//            fileChooser.setDialogTitle("Save Excel File");
//            int userSelection = fileChooser.showSaveDialog(this);
//
//            if (userSelection == JFileChooser.APPROVE_OPTION) {
//                File fileToSave = fileChooser.getSelectedFile();
//
//                // Append ".xlsx" if not present
//                if (!fileToSave.getName().toLowerCase().endsWith("hi.xlsx")) {
//                    fileToSave = new File(fileToSave.getParentFile(), fileToSave.getName() + "hi.xlsx");
//                }
//
//                // Write the workbook to the selected file
//                try (FileOutputStream fileOut = new FileOutputStream(fileToSave)) {
//                    workbook.write(fileOut);
//                }
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }


}
