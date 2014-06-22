/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package premailer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import static java.util.Collections.list;
import java.util.List;

/**
 *
 * @author bcorwin
 */
public class PRemailer {

    /**
     * @param args the command line arguments
     */
    public static String[] genEmailBody(String filename) {
        importEmailList inputFile = new importEmailList();
        try {
            inputFile.setInputFile(filename);
            String[] emailBody;
            emailBody = inputFile.readCol("Email Body", "0", false);
            return emailBody;
        } catch(IOException e) {
            throw new Error("Error reading file. May not exist or not in xls format.");
        }
    }
    
    public static Agency[] genAgencies(String filename) {
        importEmailList inputFile = new importEmailList();
        try {
            inputFile.setInputFile(filename);
            String[] agencyList, nameList, timeList, roleList, noteList;
            agencyList = inputFile.readCol("Actor List", "Agency", true);
            nameList = inputFile.readCol("Actor List", "Name", true);
            timeList = inputFile.readCol("Actor List", "Time", true);
            roleList = inputFile.readCol("Actor List", "Role", true);
            noteList = inputFile.readCol("Actor List", "Notes", false);
            
            String[] allAgencies, allAgents, allEmails;
            allAgencies = inputFile.readCol("Agency List", "Agency", true);
            String [] emailsToSend = getEmailList(allAgencies, "Agency List", agencyList);
            
            allAgents = inputFile.readCol("Agency List", "Agent Names", true);
            allEmails = inputFile.readCol("Agency List", "Agent Emails", true);
            //Build list of agencies
            Agency [] agencies = new Agency[emailsToSend.length - 1];
            for(int agnt = 1; agnt < emailsToSend.length; agnt++) {
                agencies[agnt-1] = new Agency(allAgencies[agnt], allAgents[agnt], allEmails[agnt]);
                //Build tables
                for(int row = 1; row < agencyList.length; row++) {
                    if(agencies[agnt-1].Agency.equals(agencyList[row])) {
                        agencies[agnt-1].addRow(timeList[row], nameList[row], roleList[row], noteList[row]);
                    }
                }
            }
            Arrays.sort(agencies);
            return agencies;
            
        } catch(IOException e) {
            throw new Error();
        }
    }
    public static String[] getEmailList(String[] inArray, String sheetName, String[] chkArray) {
        List<String> unique = new ArrayList<String>();
        List<String> outList = new ArrayList<String>();
        List<String> inList = Arrays.asList(chkArray);
        for(int i = 0; i < inArray.length; i++) {
            if(unique.contains(inArray[i])) {
                throw new Error("Value '" + inArray[i]
                        + "' appears more than once in '" + sheetName  + "'.");
            } else {
                unique.add(inArray[i]);
                if(inList.contains(inArray[i])) {
                    outList.add(inArray[i]);
                }
            }
        }
        String[] output = outList.toArray(new String[outList.size()]);
        return output;
    }
}
