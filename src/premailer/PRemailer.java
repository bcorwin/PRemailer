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
    static public importEmailList inputFile = new importEmailList();

    /**
     * @param args the command line arguments
     */
    public static String[] genEmailBody(String filename) {
        try {
            inputFile.setInputFile(filename);
            String[] emailBody;
            emailBody = inputFile.readCol("Email Body", "0", false);
            return emailBody;
        } catch(IOException e) {
            throw new Error("Error reading file. May not exist or not in xls format.");
        }
    }
    public static String[] getUserList(String filename) {
        try {
            inputFile.setInputFile(filename);
        } catch(IOException e) {
            throw new Error("Unknown error reading file.");
        }
        String[] output;
        output = inputFile.readCol("User List", "User Emails", false);
        return output;
    }
    
    public static Agency[] genAgencies(String filename) {
        try {
            inputFile.setInputFile(filename);
            } catch(IOException e) {
                throw new Error("Unknown error reading file.");
            }
            String[] agencyList, nameList, timeList, roleList, noteList;
            agencyList = inputFile.readCol("Actor List", "Agency", true);
            nameList = inputFile.readCol("Actor List", "Name", true);
            timeList = inputFile.readCol("Actor List", "Time", true);
            roleList = inputFile.readCol("Actor List", "Role", true);
            noteList = inputFile.readCol("Actor List", "Notes", false);
            
            String[] allAgencies, allAgents, allEmails;
            allAgencies = inputFile.readCol("Agency List", "Agency", true);
            allAgents = inputFile.readCol("Agency List", "Agent Names", true);
            allEmails = inputFile.readCol("Agency List", "Agent Emails", true);
            
            Agency [] agencies = getAgencyList(allAgencies, agencyList, allAgents, allEmails);
            //Build tables
            for(int agnt = 0; agnt < agencies.length; agnt++) {
                for(int row = 1; row < agencyList.length; row++) {
                    if(agencies[agnt].Agency.equals(agencyList[row])) {
                        agencies[agnt].addRow(timeList[row], nameList[row], 
                                roleList[row], noteList[row], row + 1);
                    }
                }
            }
            Arrays.sort(agencies);
            return agencies;
    }
    public static Agency[] getAgencyList(String[] agencyArray,
            String[] agntArray, String[] allAgents, String[] allEmails) {
        //Check that agencyList is unique
        List<String> unique = new ArrayList();
        for(int i = 1; i < agencyArray.length; i++) {
            String chkAgnt = agencyArray[i].toUpperCase().trim();
            if(unique.contains(chkAgnt)) {
                inputFile.addError("The agency '" + chkAgnt
                        + "' appears more than once in the sheet 'Agency List'."
                        + " See row #" + (i+1) + ".");
            } else {
                unique.add(chkAgnt);
            }
        }
        //Check that all values of agntArray are in agencyList
         List<String> agencyList = Arrays.asList(agencyArray);
        for(int j = 1; j < agntArray.length; j++) {
            String chkAgnt = agntArray[j];
            if(!agencyList.contains(chkAgnt)) {
                inputFile.addError("The agent '" + chkAgnt
                        + "' does not exist in the sheet 'Agency List'."
                        + " See row #" + (j+1) + ".");
            }
        }
        //Output overlapping list
        List<Agency> outList = new ArrayList();
        List<String> inList = Arrays.asList(agntArray);
        for(int k = 1; k < agencyArray.length; k++) {
            if(inList.contains(agencyArray[k])) {
                Agency addAgnt = new Agency(agencyArray[k], allAgents[k], allEmails[k]);
                outList.add(addAgnt);
            }
        }
        Agency[] output = outList.toArray(new Agency[outList.size()]);
        return output;
    }
}
