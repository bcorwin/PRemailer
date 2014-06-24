/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package premailer;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author bcorwin
 */
public class importEmailList {
    private String inputFile;
    private Workbook w;
    public String errorList = "";
    
    public void addError(String errorMsg) {
        this.errorList += errorMsg + "<br>";
    }
    
    public void setInputFile(String inputFile) throws IOException {
        
        this.inputFile = inputFile;
        File inputWorkbook = new File(inputFile);
        try {
            w = Workbook.getWorkbook(inputWorkbook);
        } catch (BiffException e) {
            e.printStackTrace();
        }
    }
    public String[] readCol(String sheetName, String col, Boolean noEmpty) {
        List<String> colList = new ArrayList<String>();
        //Check that sheet exists
        String sheetList = Arrays.toString(w.getSheetNames());
        if(sheetList.contains(sheetName) == false) {
            addError("Missing required sheet, '" + sheetName + "'.");
            throw new Error(this.errorList);
        }
        else {
            Sheet body = w.getSheet(sheetName);
            int numRows = getMaxRow(sheetName, 10);
            
            //Get col number if given a name else use that number
            int colNum = -1;
            try {
                colNum = Integer.parseInt(col);
            } catch (NumberFormatException e) {
                String chkCell;
                col = col.toUpperCase().trim();
                for (int c = 0; c < body.getColumns(); c++) {
                    chkCell = body.getCell(c,0).getContents().toUpperCase().trim();
                    if(chkCell.equals(col)) {
                        colNum = c;
                        //Known possible error. If there are more than one column names
                        break;
                    }
                }
                if(colNum == -1) {
                    addError("Column '" + col
                            + "' does not exist in sheet '" + sheetName + "'.");
                    throw new Error(this.errorList);
                }
            }

            for (int i = 0; i <= numRows; i++) {
                Cell cell = body.getCell(colNum, i);
                colList.add(cell.getContents());
                if(noEmpty == true & cell.getType() == CellType.EMPTY) {
                    addError("Column '" + col + "' is missing a value in sheet '"
                            + sheetName + "'. See row #" + (i+1) +".");
                }
            }
        }
        String[] output = colList.toArray(new String[colList.size()]);
        return output;
    }
    public int getMaxRow(String sheetName, int scanRows) {
        int totMax = 0, colMax = 0, blankCnt = 0;
        Sheet body = w.getSheet(sheetName);
        for(int col = 0 ;  col < body.getColumns(); col++) {
            colMax = 0; blankCnt = 0;
            for(int row = 0; row < body.getRows(); row++) {
                Cell currCell = body.getCell(col, row);
                if(currCell.getType() != CellType.EMPTY) {
                    colMax = row;
                    blankCnt = 0;
                } else blankCnt += 1;
                if(blankCnt >= scanRows) break;
            }
            if(colMax > totMax) totMax = colMax;
        }
        return totMax;
    }
    public void chkErrors() {
        if(!this.errorList.equals("")) throw new Error(this.errorList);
    }
    public void clrErrors() {
        this.errorList = "";
    }
}
