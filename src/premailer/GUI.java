/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package premailer;

import static java.awt.SystemColor.text;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintWriter;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.DefaultListModel;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import static premailer.PRemailer.genAgencies;
import static premailer.PRemailer.genEmailBody;
import static premailer.PRemailer.getUserList;

/**
 *
 * @author bcorwin
 */
public class GUI extends javax.swing.JFrame {
    private static String[] emailBody;
    private static Agency[] Agencies;
    private String verNum  = "0.99";
    private int currNum, numAgnts;
    private boolean isModified = false, loaded = false;
    FileNameExtensionFilter xlsFilter = new FileNameExtensionFilter("XLS files only", "xls");

    /**
     * Creates new form GUI
     */
    public GUI() {
        initComponents();
        versionLabel.setText("Version " + verNum);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        fileChooser = new javax.swing.JFileChooser();
        sendDialog = new javax.swing.JDialog();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        senderPass = new javax.swing.JPasswordField();
        jSeparator1 = new javax.swing.JSeparator();
        jLabel4 = new javax.swing.JLabel();
        ccField = new javax.swing.JTextField();
        cancelSendButton = new javax.swing.JButton();
        sendAllButton = new javax.swing.JButton();
        jSeparator2 = new javax.swing.JSeparator();
        jLabel5 = new javax.swing.JLabel();
        testEmailField = new javax.swing.JTextField();
        senderEmailBox = new javax.swing.JComboBox();
        sentDialog = new javax.swing.JDialog();
        jScrollPane2 = new javax.swing.JScrollPane();
        sendingPane = new javax.swing.JTextPane();
        jButton4 = new javax.swing.JButton();
        emailProgress = new javax.swing.JProgressBar();
        fileTextField = new javax.swing.JTextField();
        browseButton = new javax.swing.JButton();
        loadButton = new javax.swing.JButton();
        exitButton = new javax.swing.JButton();
        sendButton = new javax.swing.JButton();
        nextButton = new javax.swing.JButton();
        prevButton = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        emailTextPane = new javax.swing.JTextPane();
        emailLabel = new javax.swing.JLabel();
        versionLabel = new javax.swing.JLabel();
        jScrollPane3 = new javax.swing.JScrollPane();
        emailList = new javax.swing.JList();
        subjectField = new javax.swing.JTextField();
        toField = new javax.swing.JTextField();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        discardButton = new javax.swing.JButton();
        saveButton = new javax.swing.JButton();

        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.setApproveButtonText("Open");
        fileChooser.setApproveButtonToolTipText("");
        fileChooser.setFileFilter(xlsFilter);

        sendDialog.setTitle("Send Emails");
        sendDialog.setAlwaysOnTop(true);
        sendDialog.setMinimumSize(new java.awt.Dimension(360, 250));

        jLabel2.setText("Sender email:");

        jLabel3.setText("Sender pw:");

        jLabel4.setText("CC:");

        ccField.setToolTipText("If filled in, all emails will send a CC to this address.");

        cancelSendButton.setText("Cancel");
        cancelSendButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                cancelSendButtonActionPerformed(evt);
            }
        });

        sendAllButton.setText("Send All");
        sendAllButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                sendAllButtonActionPerformed(evt);
            }
        });

        jLabel5.setText("Test email:");
        jLabel5.setToolTipText("All emails will be sent to the test email.");

        testEmailField.setToolTipText("If filled in, all emails will be sent to this address.");

        senderEmailBox.setEditable(true);

        javax.swing.GroupLayout sendDialogLayout = new javax.swing.GroupLayout(sendDialog.getContentPane());
        sendDialog.getContentPane().setLayout(sendDialogLayout);
        sendDialogLayout.setHorizontalGroup(
            sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jSeparator1)
            .addComponent(jSeparator2, javax.swing.GroupLayout.Alignment.TRAILING)
            .addGroup(sendDialogLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(sendDialogLayout.createSequentialGroup()
                        .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel2)
                            .addComponent(jLabel3))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(senderPass)
                            .addComponent(senderEmailBox, 0, 271, Short.MAX_VALUE)))
                    .addGroup(sendDialogLayout.createSequentialGroup()
                        .addComponent(jLabel4)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(ccField))
                    .addGroup(sendDialogLayout.createSequentialGroup()
                        .addComponent(cancelSendButton)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addComponent(sendAllButton, javax.swing.GroupLayout.PREFERRED_SIZE, 81, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(sendDialogLayout.createSequentialGroup()
                        .addComponent(jLabel5)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(testEmailField)))
                .addContainerGap())
        );
        sendDialogLayout.setVerticalGroup(
            sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(sendDialogLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(senderEmailBox, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(senderPass, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 5, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4)
                    .addComponent(ccField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 10, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel5)
                    .addComponent(testEmailField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addGroup(sendDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(cancelSendButton)
                    .addComponent(sendAllButton))
                .addContainerGap())
        );

        sentDialog.setTitle("Sending emails...");
        sentDialog.setMinimumSize(new java.awt.Dimension(700, 375));

        sendingPane.setEditable(false);
        sendingPane.setContentType("text/html"); // NOI18N
        jScrollPane2.setViewportView(sendingPane);

        jButton4.setText("Exit");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout sentDialogLayout = new javax.swing.GroupLayout(sentDialog.getContentPane());
        sentDialog.getContentPane().setLayout(sentDialogLayout);
        sentDialogLayout.setHorizontalGroup(
            sentDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(sentDialogLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(sentDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, sentDialogLayout.createSequentialGroup()
                        .addGap(0, 629, Short.MAX_VALUE)
                        .addComponent(jButton4))
                    .addComponent(emailProgress, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap())
        );
        sentDialogLayout.setVerticalGroup(
            sentDialogLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(sentDialogLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(emailProgress, javax.swing.GroupLayout.DEFAULT_SIZE, 26, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 261, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton4)
                .addContainerGap())
        );

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Main Menu");
        setMinimumSize(new java.awt.Dimension(1000, 700));

        fileTextField.setText("Paste filename here or select browse. File must be in xls (old Excel) format.");
        fileTextField.setPreferredSize(new java.awt.Dimension(720, 20));
        fileTextField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileTextFieldActionPerformed(evt);
            }
        });

        browseButton.setText("Browse");
        browseButton.setMaximumSize(new java.awt.Dimension(73, 23));
        browseButton.setMinimumSize(new java.awt.Dimension(80, 23));
        browseButton.setPreferredSize(new java.awt.Dimension(73, 23));
        browseButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                browseButtonActionPerformed(evt);
            }
        });

        loadButton.setText("Load");
        loadButton.setMaximumSize(new java.awt.Dimension(80, 23));
        loadButton.setMinimumSize(new java.awt.Dimension(80, 23));
        loadButton.setPreferredSize(new java.awt.Dimension(80, 23));
        loadButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                loadButtonActionPerformed(evt);
            }
        });

        exitButton.setText("Exit");
        exitButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exitButtonActionPerformed(evt);
            }
        });

        sendButton.setText("Send");
        sendButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                sendButtonActionPerformed(evt);
            }
        });

        nextButton.setText("Next");
        nextButton.setMaximumSize(new java.awt.Dimension(85, 23));
        nextButton.setMinimumSize(new java.awt.Dimension(85, 23));
        nextButton.setPreferredSize(new java.awt.Dimension(85, 23));
        nextButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                nextButtonActionPerformed(evt);
            }
        });

        prevButton.setText("Previous");
        prevButton.setMaximumSize(new java.awt.Dimension(85, 23));
        prevButton.setMinimumSize(new java.awt.Dimension(85, 23));
        prevButton.setPreferredSize(new java.awt.Dimension(85, 23));
        prevButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                prevButtonActionPerformed(evt);
            }
        });

        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Created by Ben Corwin (bscorwin@gmail.com)");

        emailTextPane.setContentType("text/html"); // NOI18N
        emailTextPane.setFont(new java.awt.Font("Courier New", 0, 12)); // NOI18N
        emailTextPane.addCaretListener(new javax.swing.event.CaretListener() {
            public void caretUpdate(javax.swing.event.CaretEvent evt) {
                emailTextPaneCaretUpdate(evt);
            }
        });
        jScrollPane1.setViewportView(emailTextPane);

        emailLabel.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        emailLabel.setText("File not loaded.");

        versionLabel.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        versionLabel.setText("Version N.N");

        emailList.setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        emailList.addListSelectionListener(new javax.swing.event.ListSelectionListener() {
            public void valueChanged(javax.swing.event.ListSelectionEvent evt) {
                emailListValueChanged(evt);
            }
        });
        jScrollPane3.setViewportView(emailList);

        subjectField.addCaretListener(new javax.swing.event.CaretListener() {
            public void caretUpdate(javax.swing.event.CaretEvent evt) {
                subjectFieldCaretUpdate(evt);
            }
        });
        subjectField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                subjectFieldActionPerformed(evt);
            }
        });

        toField.addCaretListener(new javax.swing.event.CaretListener() {
            public void caretUpdate(javax.swing.event.CaretEvent evt) {
                toFieldCaretUpdate(evt);
            }
        });

        jLabel6.setText("Subject:");

        jLabel7.setText("To:");

        discardButton.setText("Discard");
        discardButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                discardButtonActionPerformed(evt);
            }
        });

        saveButton.setText("Save");
        saveButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(exitButton, javax.swing.GroupLayout.PREFERRED_SIZE, 73, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, 792, Short.MAX_VALUE)
                            .addComponent(versionLabel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(sendButton, javax.swing.GroupLayout.PREFERRED_SIZE, 73, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(fileTextField, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(browseButton, javax.swing.GroupLayout.PREFERRED_SIZE, 80, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(loadButton, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(emailLabel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addComponent(prevButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(nextButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jScrollPane1))
                            .addGroup(layout.createSequentialGroup()
                                .addGap(10, 10, 10)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                    .addComponent(jLabel6)
                                    .addComponent(jLabel7))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(toField)
                                    .addComponent(subjectField))
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(saveButton, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addComponent(discardButton, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 85, javax.swing.GroupLayout.PREFERRED_SIZE))))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(fileTextField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(loadButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(browseButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(subjectField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6)
                            .addComponent(discardButton)))
                    .addComponent(emailLabel, javax.swing.GroupLayout.Alignment.TRAILING))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(toField, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel7)
                    .addComponent(saveButton)
                    .addComponent(nextButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(prevButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 504, Short.MAX_VALUE)
                    .addComponent(jScrollPane3))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(exitButton)
                            .addComponent(sendButton)
                            .addComponent(versionLabel))
                        .addContainerGap())
                    .addComponent(jLabel1, javax.swing.GroupLayout.Alignment.TRAILING)))
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void browseButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_browseButtonActionPerformed
        String runPath = System.getProperty("user.dir");
        fileChooser.setCurrentDirectory(new File(runPath));
        int returnVal = fileChooser.showOpenDialog(this);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            fileTextField.setText(file.getAbsolutePath());
        }
    }//GEN-LAST:event_browseButtonActionPerformed

    private void fileTextFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_fileTextFieldActionPerformed
        loadButton.doClick();
    }//GEN-LAST:event_fileTextFieldActionPerformed

    private void exitButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exitButtonActionPerformed
        // TODO add your handling code here:
        System.exit(0);
    }//GEN-LAST:event_exitButtonActionPerformed

    private void loadButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_loadButtonActionPerformed
        String filename = fileTextField.getText();
        emailTextPane.setText("");
        isModified = false;
        try {
            PRemailer.inputFile.clrErrors();
            emailBody = genEmailBody(filename);
            Agencies = genAgencies(filename);
            numAgnts = Agencies.length;
            String[] userList = getUserList(filename);
            senderEmailBox.removeAllItems();
            for(int i = 1; i < userList.length; i++) {
                senderEmailBox.addItem(userList[i]);
            }
            PRemailer.inputFile.chkErrors();
            DefaultListModel addList = new DefaultListModel();
            for(int agnt = 0; agnt < numAgnts; agnt++) {
                Agencies[agnt].setEmail(emailBody);
                addList.addElement(Agencies[agnt].Agency);
            }
            emailList.setModel(addList);
            currNum = 0;
            emailList.setSelectedIndex(currNum);
            loaded = true;
        } catch (Error e) {
            emailTextPane.setText("<font color = \'red\'>" + e.getMessage() + "</font>");
            loaded = false;
        }
    }//GEN-LAST:event_loadButtonActionPerformed

    private void nextButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_nextButtonActionPerformed
        previewEmail(currNum+1, true);
    }//GEN-LAST:event_nextButtonActionPerformed

    private void prevButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_prevButtonActionPerformed
        previewEmail(currNum-1, true);
    }//GEN-LAST:event_prevButtonActionPerformed

    private void sendButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_sendButtonActionPerformed
        if(loaded == true & isModified == false) sendDialog.setVisible(true);
        else if(isModified == false) {
            JOptionPane.showMessageDialog(null, "Please load a file before trying to send",
                "No file loaded.", JOptionPane.ERROR_MESSAGE);
        } else if(isModified == true) {
            JOptionPane.showMessageDialog(null, "Please go back and save or discard changes.",
                "File not saved.", JOptionPane.ERROR_MESSAGE);            
        }
    }//GEN-LAST:event_sendButtonActionPerformed

    private void cancelSendButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_cancelSendButtonActionPerformed
        sendDialog.setVisible(false);
    }//GEN-LAST:event_cancelSendButtonActionPerformed

    private void sendAllButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_sendAllButtonActionPerformed
        String results = "";
        sendDialog.setVisible(false);
        sentDialog.setVisible(true);
        for(int agnt = 0; agnt < numAgnts; agnt++) {
            emailProgress.setValue((int) Math.floor((agnt+1)/numAgnts)*100);
            results += Agencies[agnt].sendEmail(senderEmailBox.getSelectedItem().toString(), senderPass.getText()
                    , ccField.getText(), testEmailField.getText());
            sendingPane.setText(results);
        }
        results += "<br><font color=\'green\'><b>COMPLETED!</b></font><br>";
        results += "<br>Notes:<br>" 
                + "If all emails failed: the sender email/password may be wrong or not gmail.<br>"
                + "If it continues to fail, try signing into your email with your web browser before returning.";
        sendingPane.setText(results);
    }//GEN-LAST:event_sendAllButtonActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        sentDialog.setVisible(false);
    }//GEN-LAST:event_jButton4ActionPerformed

    private void emailListValueChanged(javax.swing.event.ListSelectionEvent evt) {//GEN-FIRST:event_emailListValueChanged
        previewEmail(emailList.getSelectedIndex(), false);
    }//GEN-LAST:event_emailListValueChanged

    private void subjectFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_subjectFieldActionPerformed

    }//GEN-LAST:event_subjectFieldActionPerformed

    private void subjectFieldCaretUpdate(javax.swing.event.CaretEvent evt) {//GEN-FIRST:event_subjectFieldCaretUpdate
        if(subjectField.getText().equals(Agencies[currNum].Subject)) isModified = false;
        else isModified = true;
    }//GEN-LAST:event_subjectFieldCaretUpdate

    private void saveButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveButtonActionPerformed
        saveChanges();
    }//GEN-LAST:event_saveButtonActionPerformed

    private void discardButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_discardButtonActionPerformed
        isModified = false;
        previewEmail(currNum, false);
    }//GEN-LAST:event_discardButtonActionPerformed

    private void toFieldCaretUpdate(javax.swing.event.CaretEvent evt) {//GEN-FIRST:event_toFieldCaretUpdate
        if(toField.getText().equals(Agencies[currNum].Emails)) isModified = false;
        else isModified = true;
    }//GEN-LAST:event_toFieldCaretUpdate

    private void emailTextPaneCaretUpdate(javax.swing.event.CaretEvent evt) {//GEN-FIRST:event_emailTextPaneCaretUpdate
        //Known issue: if you change the number of spaces in the text, it is not detected.
        if(loaded == true) {
            String p1 = "\\s+";
            String inPane = emailTextPane.getText().replaceAll(p1,"");
            String stored = Agencies[currNum].Body.replaceAll(p1,"");
            if(!inPane.equals(stored)) isModified = true;
            else isModified = false;
        }
    }//GEN-LAST:event_emailTextPaneCaretUpdate
    public void previewEmail(int val, boolean selectNew) {
        if(isModified == false) {
            currNum = ((val % numAgnts) + numAgnts) % numAgnts;
            if(selectNew == true) emailList.setSelectedIndex(currNum);
            subjectField.setText(Agencies[currNum].Subject);
            toField.setText(Agencies[currNum].Emails);
            emailTextPane.setText(Agencies[currNum].Body);
            emailLabel.setText((currNum + 1) + " of " + numAgnts);
        } else {
            int ans = JOptionPane.showConfirmDialog(null, "Email has been modified."
                    + " Do you want to save?");
            if(ans == JOptionPane.NO_OPTION) {
                isModified = false;
                previewEmail(val, true);
            } else if(ans == JOptionPane.YES_OPTION) {
                saveChanges();
                previewEmail(val, true);
            }
        }
    }
    private void saveChanges() {
        Agency currAgnt = Agencies[currNum];
        currAgnt.Subject = currAgnt.convertCommand(subjectField.getText());
        currAgnt.Emails = toField.getText();
        currAgnt.Body = currAgnt.convertCommand(emailTextPane.getText());
        isModified = false;
    }
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new GUI().setVisible(true);
            }
        });
        //  To do:
        //  -Sort on time
        //  -"Loading screen" during the "sending" phase
        //  -Let the user decide the column order using the excel
        //  -Make the email previews editable -> store the changes -> email those changes.

    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton browseButton;
    private javax.swing.JButton cancelSendButton;
    public static javax.swing.JTextField ccField;
    private javax.swing.JButton discardButton;
    private javax.swing.JLabel emailLabel;
    private javax.swing.JList emailList;
    private javax.swing.JProgressBar emailProgress;
    private javax.swing.JTextPane emailTextPane;
    private javax.swing.JButton exitButton;
    private javax.swing.JFileChooser fileChooser;
    private javax.swing.JTextField fileTextField;
    private javax.swing.JButton jButton4;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JSeparator jSeparator1;
    private javax.swing.JSeparator jSeparator2;
    private javax.swing.JButton loadButton;
    private javax.swing.JButton nextButton;
    private javax.swing.JButton prevButton;
    private javax.swing.JButton saveButton;
    private javax.swing.JButton sendAllButton;
    private javax.swing.JButton sendButton;
    private javax.swing.JDialog sendDialog;
    public static javax.swing.JComboBox senderEmailBox;
    public static javax.swing.JPasswordField senderPass;
    public static javax.swing.JTextPane sendingPane;
    private javax.swing.JDialog sentDialog;
    private javax.swing.JTextField subjectField;
    public static javax.swing.JTextField testEmailField;
    private javax.swing.JTextField toField;
    private javax.swing.JLabel versionLabel;
    // End of variables declaration//GEN-END:variables
}
