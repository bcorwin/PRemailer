/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package premailer;

import java.text.SimpleDateFormat;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.mail.*;
import javax.mail.internet.*;

/**
 *
 * @author bcorwin
 */
public class Agency implements Comparable<Agency>{
    public String Agency, Names, Emails, Subject, Body;
    public String[][] Table = new String[1][5];
    
    public Agency (String setAgency, String setNames, String setEmails) {
        Agency = setAgency;
        Names = setNames;
        Emails = setEmails;
        
        Table[0][0] = "Time";
        Table[0][1] = "Name";
        Table[0][2] = "Agency";
        Table[0][3] = "Role";
        Table[0][4] = "Notes";
    }

    Agency() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    public void addRow(String time, String name, String role, String note, int rowID) {
        int maxRow = this.Table.length, newRow = maxRow;
        SimpleDateFormat sdf = new SimpleDateFormat("hh:mm a");
        try {
            Date addTime = sdf.parse(time);
            for(int row = 1; row < maxRow; row++) {
                try {
                    Date currTime = sdf.parse(this.Table[row][0]);
                    if(addTime.before(currTime)) {
                        newRow = row;
                        break;
                    }
                } catch (java.text.ParseException e1) {}
            }
        } catch (java.text.ParseException e) {
            PRemailer.inputFile.addError("The formating for the column 'Time' may be incorrect."
                    + " It should be in the form 'HH:MM AM/PM'." + 
                    " See row #" + rowID + ".");
        }
        String[][] newTable = new String[maxRow + 1][5];
        String[][] oldTable = this.Table;
        for(int row = 0; row <= maxRow; row++) {
            if(row < newRow) {
                newTable[row] = oldTable[row];
            } else if (row == newRow) {
                newTable[row][0] = time;
                newTable[row][1] = name;
                newTable[row][2] = this.Agency;
                newTable[row][3] = role;
                newTable[row][4] = note;
            } else {
                newTable[row] = oldTable[row-1];
            }
            this.Table= newTable;

        }
    }
    
    @Override
    public int compareTo(Agency agnt) {
        return this.Agency.compareTo(agnt.Agency);
    }
    
    public String createHTMLtable() {
        String output = "<table border = \"1\" align = \"center\" style = \"width:500px\">";
        for(int rowNum = 0; rowNum < this.Table.length; rowNum++) {
            String[] row = this.Table[rowNum];
            output = output + "<tr>";
            for(int col = 0; col < row.length; col++) {
                if(rowNum == 0) {
                    output = output + "<th align = \"left\">" + row[col] + "</th>";
                } else {
                    output = output + "<td align = \"left\">" + row[col] + "</td>";
                }
            }
            output = output + "</tr>";
        }
        output = output + "</table>";
        return output;
    }

    public void setEmail(String[] emailBody) {
        String newLine, subject = "";
        String body = "<html><head></head><body>";
        for(int line = 0; line < emailBody.length; line++) {
            newLine = this.convertCommand(emailBody[line]);
            if(line == 0) {
                subject = newLine;
            } else {
                body = body + newLine + "<br>";
            }
        }
        body = body + "</body></html>";
        this.Body = body;
        this.Subject = subject;
    }
    public String convertCommand(String input) {
        String output = input;
        output = output.replaceAll("(\\<|&lt;)(?i)LIST(\\>|&gt;)", this.createHTMLtable());
        output = output.replaceAll("(\\<|&lt;)(?i)AGNTNAME(\\>|&gt;)", this.Names);
        return output;
    }
    
    public String sendEmail(String from, String pass, String cc, String testEmail) {
        String output = "";
        String[] to, ccList;
        if("".equals(testEmail)) to = this.Emails.split(",");
            else to = testEmail.split(",");
        ccList = cc.split(",");

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        Session session = Session.getInstance(props,
            new javax.mail.Authenticator() {
                protected PasswordAuthentication getPasswordAuthentication() {
                    return new PasswordAuthentication(from, pass);
                }
            }
        );
        
        try {
            MimeMessage message = new MimeMessage(session);
            message.setFrom(new InternetAddress(from));
            for(String currTo: to) {
                message.addRecipient(Message.RecipientType.TO, new InternetAddress(currTo));
            }
            if(!"".equals(cc)) {
                for(String currCC: ccList) {
                    message.addRecipient(Message.RecipientType.CC, new InternetAddress(currCC));
                }
            }
            message.setSubject(this.Subject);
            message.setContent(this.Body, "text/html");
            Transport.send(message);
            output += "Message successfully sent to " + this.Agency
                    + " (" + Arrays.toString(to) + ").<br>";
        } catch (MessagingException mex) {
//            mex.printStackTrace();
            output += "<font color = \"red\">Email failed to send for "
                    + this.Agency + " (" + Arrays.toString(to) + ").</font><br>";
      }
        return output;
   }
}
