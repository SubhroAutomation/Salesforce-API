package Utilities;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.LinkedHashSet;
import java.util.Iterator;
import java.util.Set;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

public class ExcelPOI {
	
	SubhroAutomation ......
	
    //public static int iLogSheetRowCounter=1;
    public static final int EXCELWRITE_SUCCESS = -2;
    public static final int EXCELWRITE_FAILURE = -3;
    public static String strTestDataFilePath = "";

    public static void browseForExcelFile(){
    	// Browse for the Excel file
    	
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx");
        fileChooser.setFileFilter(filter);

        fileChooser.setDialogTitle("Select the RTP sheet");
        int userSelection = fileChooser.showDialog(null, "Select Excel");

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fExcelFile = fileChooser.getSelectedFile();
            strTestDataFilePath = fExcelFile.getAbsolutePath();
            System.out.println("Excel file: " + fExcelFile.getAbsolutePath());
        }
    }
    
    public static int CheckifFileOpen(String strExcelWorkbook){
       try {
            //FileOutputStream testOut = new FileOutputStream(strExcelWorkbook, );
            FileWriter testOut = new FileWriter(strExcelWorkbook, true);
            testOut.close();
            return EXCELWRITE_SUCCESS;
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,"The Excel file: " + strExcelWorkbook + " is already open.  Please close it before continuing..","File Open Error",JOptionPane.ERROR_MESSAGE); 		
 
            System.out.println("The Excel file: "
                            + strExcelWorkbook
                            + " is already open.  Please close it before continuing..");
            return EXCELWRITE_FAILURE;
        } 
    }

    public static boolean CreateNewExcelFile(String strExcelWorkbookName){
		Workbook wbTestDataExcelWB = null;
		
		try {
		    String strFileExtnsn = FilenameUtils.getExtension(strExcelWorkbookName);
		    //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);
		
		    if(strFileExtnsn.equals("xlsx")){
		        wbTestDataExcelWB = new XSSFWorkbook();
		    }
		    else if(strFileExtnsn.equals("xls")){
		        wbTestDataExcelWB = new HSSFWorkbook();
		    }
		    
	   	    FileOutputStream fileOut;
	   	    fileOut = new FileOutputStream(strExcelWorkbookName);
	   	    wbTestDataExcelWB.write(fileOut);
	    	fileOut.close();
	    	
	    	return true;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return false;
		}
    }
    
    public static ArrayList<String> GetUniqueRowsfromColumn(String strSheetName, String strColNameToRead){
            ArrayList<String> alSubCapabilities =new ArrayList<String>();
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int i, j;
                boolean bColNameFound = false;
                Row row = shtTestDataSheet.getRow(0);
                for (j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bColNameFound){
                    //System.out.println(shtTestDataSheet.getLastRowNum());
                    for (i = 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                                row = shtTestDataSheet.getRow(i);
                                //System.out.println(row.getCell(j).getStringCellValue());
                                alSubCapabilities.add(row.getCell(j).getStringCellValue());
                    }
                }
                else{
                    //System.out.println("Column Name not found");
                }

                    Set<String> hs = new LinkedHashSet<>();
                    hs.addAll(alSubCapabilities);
                    alSubCapabilities.clear();
                    alSubCapabilities.addAll(hs);
            }
            catch(Exception e){
                    e.printStackTrace();
            }

            return alSubCapabilities;
    }
    
    public static ArrayList<String> GetUniqueRowsfromColumn(String strSheetName, String strColNameToRead, int iRowToStartReadingFrom){
            ArrayList<String> alSubCapabilities =new ArrayList<String>();
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int i, j;
                boolean bColNameFound = false;
                Row row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
                for (j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bColNameFound){
                    //System.out.println(shtTestDataSheet.getLastRowNum());
                    for (i = iRowToStartReadingFrom+1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                        row = shtTestDataSheet.getRow(i);
                        //System.out.println(row.getCell(j).getStringCellValue());
                        
                        try{
                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
                        		String strCellValue = row.getCell(j).getStringCellValue().trim();
                        		if (!strCellValue.equals(""))
                                    alSubCapabilities.add(strCellValue);                       		
                        	}
                        }
                        catch(Exception e2){
                            System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                        }
                        
                    }
                }
                else{
                    //System.out.println("Column Name not found");
                }

                    Set<String> hs = new LinkedHashSet<>();
                    hs.addAll(alSubCapabilities);
                    alSubCapabilities.clear();
                    alSubCapabilities.addAll(hs);
            }
            catch(Exception e){
                    e.printStackTrace();
            }

            return alSubCapabilities;
    }
    
    public static ArrayList<String> GetUniqueRowsfromColumn(String strSheetName, String strColNameToRead, int iRowToStartReadingFrom, int iEndRow){
        ArrayList<String> alSubCapabilities =new ArrayList<String>();
        try{
            //Create a object of File class to open xlsx file
            File file = new File(strTestDataFilePath);

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            Workbook wbTestDataExcelWB = null;
            String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
            //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

            if(strFileExtnsn.equals("xlsx")){
                wbTestDataExcelWB = new XSSFWorkbook(inputStream);
            }
            else if(strFileExtnsn.equals("xls")){
                wbTestDataExcelWB = new HSSFWorkbook(inputStream);
            }

            //Read sheet inside the workbook by its name
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

            int i, j;
            boolean bColNameFound = false;
            Row row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (j = 0; j <= row.getLastCellNum(); j++) {
                if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                        bColNameFound = true;
                        break;
                }
            }

            if (bColNameFound){
                //System.out.println(shtTestDataSheet.getLastRowNum());
                for (i = iRowToStartReadingFrom+1; i <= iEndRow; i++) {
                    row = shtTestDataSheet.getRow(i);
                    //System.out.println(row.getCell(j).getStringCellValue());
                    
                    try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
                    		String strCellValue = row.getCell(j).getStringCellValue().trim();
                    		if (!strCellValue.equals(""))
                                alSubCapabilities.add(strCellValue);                       		
                    	}
                    }
                    catch(Exception e2){
                        System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                    }
                    
                }
            }
            else{
                //System.out.println("Column Name not found");
            }

                Set<String> hs = new LinkedHashSet<>();
                hs.addAll(alSubCapabilities);
                alSubCapabilities.clear();
                alSubCapabilities.addAll(hs);
        }
        catch(Exception e){
                e.printStackTrace();
        }

        return alSubCapabilities;
    }
    
    public static ArrayList<String> GetAllRowsfromColumn(String strSheetName, String strColNameToRead){
            ArrayList<String> alAllRows =new ArrayList<String>();
            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int i, j;
                boolean bColNameFound = false;
                Row row = shtTestDataSheet.getRow(0);
                for (j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bColNameFound){
                    //System.out.println(shtTestDataSheet.getLastRowNum());
                    for (i = 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                    	row = shtTestDataSheet.getRow(i);
                    	
                    	try{
                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
                        		String strCellValue = row.getCell(j).getStringCellValue().trim();
                        		if (!strCellValue.equals("")){
                        			//System.out.println(row.getCell(j).getStringCellValue());
                                    alAllRows.add(row.getCell(j).getStringCellValue());  
                        		}
                        	}
                        }
                        catch(Exception e2){
                            System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                        }
                    }
                }
                else{
                    //System.out.println("Column Name not found");
                }
            }
            catch(Exception e){
                    e.printStackTrace();
            }

            return alAllRows;
    }

    public static ArrayList<String> GetAllRowsfromColumnfrmGivenRowIndex(String strSheetName, String strColNameToRead, int iStartRow){
        ArrayList<String> alAllRows =new ArrayList<String>();
        try{
                //Create a object of File class to open xlsx file
            File file = new File(strTestDataFilePath);

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            Workbook wbTestDataExcelWB = null;
            String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
            //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

            if(strFileExtnsn.equals("xlsx")){
                wbTestDataExcelWB = new XSSFWorkbook(inputStream);
            }
            else if(strFileExtnsn.equals("xls")){
                wbTestDataExcelWB = new HSSFWorkbook(inputStream);
            }

            //Read sheet inside the workbook by its name
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

            int i, j;
            boolean bColNameFound = false;
            Row row = shtTestDataSheet.getRow(iStartRow);
            for (j = 0; j <= row.getLastCellNum(); j++) {
                if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                        bColNameFound = true;
                        break;
                }
            }

            if (bColNameFound){
                //System.out.println(shtTestDataSheet.getLastRowNum());
                for (i = iStartRow + 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
                    		String strCellValue = row.getCell(j).getStringCellValue().trim();
                    		if (!strCellValue.equals("")){
                    			//System.out.println(row.getCell(j).getStringCellValue());
                                alAllRows.add(row.getCell(j).getStringCellValue());  
                    		}
                    	}
                    }
                    catch(Exception e2){
                        System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                    }
                }
            }
            else{
                //System.out.println("Column Name not found");
            }
        }
        catch(Exception e){
                e.printStackTrace();
        }

        return alAllRows;
    }
    
    public static ArrayList<String> GetAllColumnsfromRow(String strSheetName, String strRowNameToRead){
            ArrayList<String> alAllColumns =new ArrayList<String>();
            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);
                
                int iIndexofRowToRead;
                boolean bRowNameFound = false;

                //Find index of RowName
                Row row;
                for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                    row = shtTestDataSheet.getRow(iIndexofRowToRead);
                    //System.out.println(row.getCell(0).getStringCellValue() + " = " + strRowNameToRead);
                    if (row.getCell(0).getStringCellValue().equals(strRowNameToRead)){
                            bRowNameFound = true;
                            break;
                    }
                }
                
                int j;
                if (bRowNameFound){
                    row = shtTestDataSheet.getRow(iIndexofRowToRead);
                    for (j = 1; j <= row.getLastCellNum(); j++) {
                        String strCellValueRetrieved = "";
                        try{
                            strCellValueRetrieved = row.getCell(j).getStringCellValue();                            
                        }
                        catch(Exception e2){
                            strCellValueRetrieved = "";
                        }
                        if (!strCellValueRetrieved.equals(""))
                            alAllColumns.add(strCellValueRetrieved);
                    }
                }
                else{
                    //System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name not found","Excel Error",JOptionPane.ERROR_MESSAGE);
                }                
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE);
            }

            return alAllColumns;
    }
    
    public static ArrayList<String> GetSheetsfromWorkbook(){
            ArrayList<String> alSheetNames =new ArrayList<String>();
            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                    for (int i = 0; i < wbTestDataExcelWB.getNumberOfSheets(); i++)
                    {
                            alSheetNames.add(wbTestDataExcelWB.getSheetAt(i).getSheetName());
                    }
            }
            catch(Exception e){
                    e.printStackTrace();
            }

            return alSheetNames;
    }

    public static String ReadDataFromExcel(String strSheetName, String strColNameToRead, int iRowIndexToRead) throws IOException {
            String strValueRetrieved = "";

            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int i, j;
                boolean bColNameFound = false;

                Row row = shtTestDataSheet.getRow(0);
                for (j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j).getStringCellValue().equals(strColNameToRead)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bColNameFound){
                        row = shtTestDataSheet.getRow(iRowIndexToRead);
                        //System.out.println(i +" " +j);
                        //System.out.println("Value retrived: " +strValueRetrieved);
                        try{
                            strValueRetrieved = row.getCell(j).getStringCellValue();                            
                        }
                        catch(Exception e2){
                            strValueRetrieved = "";
                        }
                }
                else{
                    System.out.println("Column Name not found");
                }
            }
            catch(IOException e){
                    e.printStackTrace();
            }

            return strValueRetrieved;
    }

    public static String ReadDataFromExcel(String strSheetName, String strColNameToRead, String strRowNameToRead) throws IOException {
            String strValueRetrieved = "";
            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iIndexofRowToRead, iIndexofColToRead;
                boolean bRowNameFound = false, bColNameFound = false;

                //Find index of RowName
                Row row;
                for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                    row = shtTestDataSheet.getRow(iIndexofRowToRead);
                    //System.out.println(row.getCell(0).getStringCellValue() + " = " + strRowNameToRead);
                    if (row.getCell(0).getStringCellValue().equals(strRowNameToRead)){
                            bRowNameFound = true;
                            break;
                    }
                }

                //Find index of ColumnName
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColToRead = 0; iIndexofColToRead <= row.getLastCellNum(); iIndexofColToRead++) {
                    if (row.getCell(iIndexofColToRead).getStringCellValue().equals(strColNameToRead)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bRowNameFound && bColNameFound){
                        row = shtTestDataSheet.getRow(iIndexofRowToRead);
                        //System.out.println(i +" " +j);
                        //System.out.println("Value retrived: " +strValueRetrieved);
                        try{
                        	if (Cell.CELL_TYPE_BLANK == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = "";
                        	else if (Cell.CELL_TYPE_NUMERIC == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = Integer.toString((int)row.getCell(iIndexofColToRead).getNumericCellValue());
                        	else if (Cell.CELL_TYPE_STRING == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = row.getCell(iIndexofColToRead).getStringCellValue();                            
                        }
                        catch(Exception e2){
                            strValueRetrieved = "";
                        }                        
                }
                else{
                    //System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name not found","Excel Error",JOptionPane.ERROR_MESSAGE);
                }
            }
            catch(IOException e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE);
            }

            return strValueRetrieved;
    }

    public static String ReadDataFromExcel(String strSheetName, String strColNameXToRead, String strColNameYwhrRecToSrch, String strRecordToSrch) throws IOException {
            String strValueRetrieved = "";
            try{
                    //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iIndexofRowToRead, iIndexofColToRead;
                boolean bRowNameFound = false, bColNameXFound = false, bColNameYFound = false;

                Row row;
                
                //Find index of ColumnName
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColToRead = 0; iIndexofColToRead <= row.getLastCellNum(); iIndexofColToRead++) {
                    if (row.getCell(iIndexofColToRead).getStringCellValue().equals(strColNameYwhrRecToSrch)){
                            bColNameYFound = true;
                            break;
                    }
                }
                
                //Find index of RowName                
                for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                    row = shtTestDataSheet.getRow(iIndexofRowToRead);
                    //System.out.println(row.getCell(0).getStringCellValue() + " = " + strRowNameToRead);
                    if (row.getCell(iIndexofColToRead).getStringCellValue().equals(strRecordToSrch)){
                            bRowNameFound = true;
                            break;
                    }
                }

                //Find index of ColumnNameX
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColToRead = 0; iIndexofColToRead <= row.getLastCellNum(); iIndexofColToRead++) {
                    if (row.getCell(iIndexofColToRead).getStringCellValue().equals(strColNameXToRead)){
                            bColNameXFound = true;
                            break;
                    }
                }

                if (bRowNameFound && bColNameXFound){
                        row = shtTestDataSheet.getRow(iIndexofRowToRead);
                        //System.out.println(i +" " +j);
                        //System.out.println("Value retrived: " +strValueRetrieved);
                        try{
                        	if (Cell.CELL_TYPE_BLANK == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = "";
                        	else if (Cell.CELL_TYPE_NUMERIC == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = Integer.toString((int)row.getCell(iIndexofColToRead).getNumericCellValue());
                        	else if (Cell.CELL_TYPE_STRING == row.getCell(iIndexofColToRead).getCellType())
                        		strValueRetrieved = row.getCell(iIndexofColToRead).getStringCellValue();                            
                        }
                        catch(Exception e2){
                            strValueRetrieved = "";
                        }                        
                }
                else{
                    //System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name or Record not found","Excel Error",JOptionPane.ERROR_MESSAGE);
                }
            }
            catch(IOException e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE);
            }

            return strValueRetrieved;
    }
    
    public static void WriteDataToExcel(String strSheetName, String strColNameToWrite, int iRowIndexToWrite, String strValueToWrite) throws Exception {
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int i, j;
                boolean bColNameFound = false;

                Row row = shtTestDataSheet.getRow(0);
                for (j = 0; j <= row.getLastCellNum(); j++) {
                    if (row.getCell(j).getStringCellValue().equals(strColNameToWrite)){
                            bColNameFound = true;
                            break;
                    }
                }

                if (bColNameFound){
                    row = shtTestDataSheet.getRow(iRowIndexToWrite);
                    //System.out.println("j: " +j);

                	if (row == null)	
                		row = shtTestDataSheet.createRow(iRowIndexToWrite);
                	
                    Cell cell = row.createCell(j);
                    
                    if (Cell.CELL_TYPE_BLANK == cell.getCellType()){
                	   cell.setCellType(Cell.CELL_TYPE_STRING);
                	   cell.setCellValue(strValueToWrite);
                	} 
                }
                else{
                    System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name not found","Excel Error",JOptionPane.ERROR_MESSAGE); 
                }

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    //JOptionPane.showMessageDialog(null,"Please close the excel file first","File Open Error",JOptionPane.ERROR_MESSAGE); 
            }
    }

    public static void WriteDataToExcel(String strSheetName, String strColNameToWrite, String strRecToSearch, String strColToSearch, String strValueToWrite) throws Exception {
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iIndexofRowToWrite, iIndexofColToWrite, iIndexofColToSearch;
                boolean bRowNameFound = false, bColNameToWriteFound = false, bColNameToSearchFound=false;

                //Find index of RowName
                Row row;

              //Find index of ColumnName To Search
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColToSearch = 0; iIndexofColToSearch <= row.getLastCellNum(); iIndexofColToSearch++) {
                    if (row.getCell(iIndexofColToSearch).getStringCellValue().equals(strColToSearch)){
                            bColNameToSearchFound = true;
                            break;
                    }
                }

        //Find index of ColumnName To Write
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColToWrite = 0; iIndexofColToWrite <= row.getLastCellNum(); iIndexofColToWrite++) {
                    if (row.getCell(iIndexofColToWrite).getStringCellValue().equals(strColNameToWrite)){
                            bColNameToWriteFound = true;
                            break;
                    }
                }

                if (bColNameToSearchFound && bColNameToWriteFound){
            for (iIndexofRowToWrite = 0; iIndexofRowToWrite <= shtTestDataSheet.getLastRowNum(); iIndexofRowToWrite++) {
                row = shtTestDataSheet.getRow(iIndexofRowToWrite);
                String strCellValue = row.getCell(iIndexofColToSearch).getStringCellValue();

                //System.out.println( strCellValue + " = " + strRecordinColumnX);
                if (strCellValue.equals(strRecToSearch)){
                        bRowNameFound = true;
                        Cell cell = row.createCell(iIndexofColToWrite);
                                cell.setCellValue(strValueToWrite);
                 }
            }
            if(!bRowNameFound){
                System.out.println("Record not found in Column: " +strRecToSearch);
                JOptionPane.showMessageDialog(null,"Record not found in Column","Excel Error",JOptionPane.ERROR_MESSAGE);		
                    }
                }
                else{
                    System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name not found","Excel Error",JOptionPane.ERROR_MESSAGE);		
            }

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE); 		
            }
    }

    public static void AddNewRowtoSheet(String strSheetName, String strCapability, String strSubCapability, String strSQLName, String strFunctionality, String strSQLQuery, String strProject, int... iCounter) throws Exception {
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iSheetRowCounter;
                if (iCounter.length > 0)
                    iSheetRowCounter = iCounter[0];
                else
                    iSheetRowCounter = shtTestDataSheet.getLastRowNum() + 1;

                Row row = shtTestDataSheet.createRow(iSheetRowCounter);
                Cell cell = row.createCell(0);
                cell.setCellValue(strCapability);
                cell = row.createCell(1);
                cell.setCellValue(strSubCapability);
                cell = row.createCell(2);
                cell.setCellValue(strSQLName);
                cell = row.createCell(3);
                cell.setCellValue(strFunctionality);
                cell = row.createCell(4);
                cell.setCellValue(strSQLQuery);
                cell = row.createCell(5);
                cell.setCellValue(strProject);

                    /*Calendar cal = Calendar.getInstance();
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
            String strTime = sdf.format(cal.getTime());    
            cell = row.createCell(4);
                    cell.setCellValue(strTime);*/

                    //iSheetRowCounter ++;

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Please close the excel file first","File Open Error",JOptionPane.ERROR_MESSAGE); 		
            }
    }   

    public static void AddNewRowtoNewGrpSheet(String strSheetName, String strGroupName) throws Exception {
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iSheetRowCounter = shtTestDataSheet.getLastRowNum() + 1;

                Row row = shtTestDataSheet.createRow(iSheetRowCounter);
                Cell cell = row.createCell(0);
                cell.setCellValue(strGroupName);

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in excel operation","File Open Error",JOptionPane.ERROR_MESSAGE); 		
            }
    }        

    public static int AddNewSheet(String strSheetName) throws Exception {
            int iIndex=0;
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                iIndex = wbTestDataExcelWB.getSheetIndex(strSheetName);
                if(iIndex == -1)
                        wbTestDataExcelWB.createSheet(strSheetName);
                else
                    System.out.println("Sheet exists");

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Please close the excel file first","File Open Error",JOptionPane.ERROR_MESSAGE); 		
            }
            return iIndex;
    }

    public static ArrayList<String> GetAllRowsfromColumnYforAllOccurencesofRecordinColumnX(String strSheetName, String strColumnX, String strColumnY, String strRecordinColumnX){
        ArrayList<String> alValuesinColY = new ArrayList<String>();
        try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read sheet inside the workbook by its name
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iIndexofRowToRead, iIndexofColXToRead, iIndexofColYToRead;
                boolean bRowNameFound = false, bColNameXFound = false, bColNameYFound = false;

                Row row;

                //Find index of ColumnNameX
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColXToRead = 0; iIndexofColXToRead <= row.getLastCellNum(); iIndexofColXToRead++) {
                    if (row.getCell(iIndexofColXToRead).getStringCellValue().equals(strColumnX)){
                            bColNameXFound = true;
                            break;
                    }
                }
                //Find index of ColumnNameY
                row = shtTestDataSheet.getRow(0);
                for (iIndexofColYToRead = 0; iIndexofColYToRead <= row.getLastCellNum(); iIndexofColYToRead++) {
                    if (row.getCell(iIndexofColYToRead).getStringCellValue().equals(strColumnY)){
                            bColNameYFound = true;
                            break;
                    }
                }

                if (bColNameXFound && bColNameYFound){
                        for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                            row = shtTestDataSheet.getRow(iIndexofRowToRead);
                            String strCellValue = row.getCell(iIndexofColXToRead).getStringCellValue();

                            //System.out.println( strCellValue + " = " + strRecordinColumnX);
                            if (strCellValue.equals(strRecordinColumnX)){
                                    bRowNameFound = true;
                                    String strCellValueRetrieved = "";
                                    try{
                                        strCellValueRetrieved = row.getCell(iIndexofColYToRead).getStringCellValue();
                                    }
                                    catch(Exception e2){
                                        strCellValueRetrieved = "";
                                    }                                        /*if(row.getCell(iIndexofColYToRead).getStringCellValue().isEmpty())
                                        strCellValueRetrieved = "";
                                    else
                                        strCellValueRetrieved = row.getCell(iIndexofColYToRead).getStringCellValue();*/

                                    alValuesinColY.add(strCellValueRetrieved);
                                    //System.out.println("Value retrived: " +strCellValueRetrieved);
                            }
                        }
                        if(!bRowNameFound){
                            System.out.println("Record not found in Column: " +strRecordinColumnX);
                            JOptionPane.showMessageDialog(null,"Record not found in Column: " +strRecordinColumnX,"Excel Error",JOptionPane.ERROR_MESSAGE);
                        }
                }
                else{
                    //System.out.println("Column Name not found");
                    JOptionPane.showMessageDialog(null,"Column Name not found","Excel Error",JOptionPane.ERROR_MESSAGE);
                }
        }
        catch(IOException e){
                e.printStackTrace();
        }
        return alValuesinColY;
    }

    public static int GetRowIndex (String strSheetName, String strSubCapability, String strFunctionality, String strSQLQuery) throws IOException {		
        int iIndexofRowToRead = -1;
        try{
                //Create a object of File class to open xlsx file
            File file = new File(strTestDataFilePath);

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            Workbook wbTestDataExcelWB = null;
            String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
            //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

            if(strFileExtnsn.equals("xlsx")){
                wbTestDataExcelWB = new XSSFWorkbook(inputStream);
            }
            else if(strFileExtnsn.equals("xls")){
                wbTestDataExcelWB = new HSSFWorkbook(inputStream);
            }

            //Read sheet inside the workbook by its name
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

            boolean bRowNameFound = false;

            //Find index of RowName
            Row row;
            for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                row = shtTestDataSheet.getRow(iIndexofRowToRead);
                //System.out.println(row.getCell(0).getStringCellValue() + " = " + strRowNameToRead);
                if (row.getCell(3).getStringCellValue().equals(strSQLQuery)){
                    if (row.getCell(1).getStringCellValue().equals(strSubCapability)){
                        if (row.getCell(2).getStringCellValue().equals(strFunctionality)){
                            bRowNameFound = true;
                            break;
                        }
                    }
                }
            }		       

            if (!bRowNameFound){			    
                System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Name not found","Error",JOptionPane.ERROR_MESSAGE); 		
            }
        }
        catch(IOException e){
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,"Excel Error","Error",JOptionPane.ERROR_MESSAGE);
        }

        return iIndexofRowToRead;
    }
    
    public static void WriteDataToGroupSheet(String strSheetName, String strGrpNameToSearch, String strValueToWrite) throws Exception {
        try{
            //Create a object of File class to open xlsx file
            File file = new File(strTestDataFilePath);

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            Workbook wbTestDataExcelWB = null;
            String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
            //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

            if(strFileExtnsn.equals("xlsx")){
                wbTestDataExcelWB = new XSSFWorkbook(inputStream);
            }
            else if(strFileExtnsn.equals("xls")){
                wbTestDataExcelWB = new HSSFWorkbook(inputStream);
            }

            //Read excel sheet by sheet name    
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

            int iIndexofRowToRead, iIndexofColToWrite;
            boolean bRowNameFound = false;

            //Find index of RowName
            Row row;            
            for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                row = shtTestDataSheet.getRow(iIndexofRowToRead);
                //System.out.println(row.getCell(0).getStringCellValue() + " = " + strRowNameToRead);
                if (row.getCell(0).getStringCellValue().equals(strGrpNameToSearch)){
                        bRowNameFound = true;
                        break;
                }
            }

            //Find index of ColumnName To Write
            row = shtTestDataSheet.getRow(iIndexofRowToRead);
            iIndexofColToWrite = row.getLastCellNum();   
            
            if (bRowNameFound){
                Cell cell = row.createCell(iIndexofColToWrite);
                cell.setCellValue(strValueToWrite);            
            }
            else{
                System.out.println("Group Name not found");
                JOptionPane.showMessageDialog(null,"Group Name not found","Error",JOptionPane.ERROR_MESSAGE);		
            }

            //Close input stream
            inputStream.close();

            //Create an object of FileOutputStream class to create write data in excel file
            FileOutputStream outputStream = new FileOutputStream(file);

            //write data in the excel file
            wbTestDataExcelWB.write(outputStream);

            //close output stream
            outputStream.close();
        }
        catch(Exception e){
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,"Please close the excel file first","Excel Error",JOptionPane.ERROR_MESSAGE); 		
        }
    }
    
    public static void AddNewRowtoNewSQLConSheet(String strSheetName, String strConnName, String strUsername, String strPassword, String strHostname, String strPortNo, String strService) throws Exception {
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iSheetRowCounter = shtTestDataSheet.getLastRowNum() + 1;

                Row row = shtTestDataSheet.createRow(iSheetRowCounter);
                Cell cell = row.createCell(0);
                cell.setCellValue(strConnName);
                cell = row.createCell(1);
                cell.setCellValue(strUsername);
                cell = row.createCell(2);
                cell.setCellValue(strPassword);
                cell = row.createCell(3);
                cell.setCellValue(strHostname);
                cell = row.createCell(4);
                cell.setCellValue(strPortNo);
                cell = row.createCell(5);
                cell.setCellValue(strService);

                //Close input stream
                inputStream.close();

                //Create an object of FileOutputStream class to create write data in excel file
                FileOutputStream outputStream = new FileOutputStream(file);

                //write data in the excel file
                wbTestDataExcelWB.write(outputStream);

                //close output stream
                outputStream.close();
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE); 		
            }
    }  
    
    //Gets the Index of the Last Column in a Row (RowName passed as parameter. RowName = Value of first col in the Row)
    public static int GetLastColumnIndexinaRow(String strSheetName, String strRowName) throws Exception {
            int iLastColumnIndex = -1;
            try{
                //Create a object of File class to open xlsx file
                File file = new File(strTestDataFilePath);

                //Create an object of FileInputStream class to read excel file
                FileInputStream inputStream = new FileInputStream(file);

                Workbook wbTestDataExcelWB = null;
                String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
                //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

                if(strFileExtnsn.equals("xlsx")){
                    wbTestDataExcelWB = new XSSFWorkbook(inputStream);
                }
                else if(strFileExtnsn.equals("xls")){
                    wbTestDataExcelWB = new HSSFWorkbook(inputStream);
                }

                //Read excel sheet by sheet name    
                Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

                int iIndexofRowToRead;
                boolean bRowNameFound = false;
                
                Row row=null;
                for (iIndexofRowToRead = 0; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                    row = shtTestDataSheet.getRow(iIndexofRowToRead);
                    //System.out.println("Row Index: " +iIndexofRowToRead + ".. Value: " + row.getCell(0).getStringCellValue());
                    
                    try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		if (row.getCell(0).getStringCellValue().equals(strRowName)){
                                bRowNameFound = true;
                                iLastColumnIndex = row.getLastCellNum();
                                break;
                    		}
                    	}
                    }
                    catch(Exception e2){
                        System.out.println("In Catch for Row No: " +iIndexofRowToRead + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                    }
                    
                }
                
                if (!bRowNameFound)
                    JOptionPane.showMessageDialog(null,"Row not found","Excel Error",JOptionPane.ERROR_MESSAGE); 
                
                //Close input stream
                inputStream.close();                
            }
            catch(Exception e){
                    e.printStackTrace();
                    JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE); 		
                    iLastColumnIndex = -1;
            }
            return iLastColumnIndex;
    }
    
    //Searches a column (Index passed as parameter) for a value.. iIndexofStartRowToRead: If a column requires to be read from a particular row index, pass 0 if needed to be read from first row 
    public static int GetRowIndexofValueinCol (String strSheetName, String strRowName, int iColIndex, int iIndexofStartRowToRead) throws Exception {
        int iRowIndex = -1;
        try{
            //Create a object of File class to open xlsx file
            File file = new File(strTestDataFilePath);

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            Workbook wbTestDataExcelWB = null;
            String strFileExtnsn = FilenameUtils.getExtension(strTestDataFilePath);
            //System.out.println("Excel File: " +ConfigFile.strTestDataFilePath + " Extnsn: " +strFileExtnsn);

            if(strFileExtnsn.equals("xlsx")){
                wbTestDataExcelWB = new XSSFWorkbook(inputStream);
            }
            else if(strFileExtnsn.equals("xls")){
                wbTestDataExcelWB = new HSSFWorkbook(inputStream);
            }

            //Read excel sheet by sheet name    
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetName);

            int iIndexofRowToRead;
            boolean bRowNameFound = false;
            
            Row row=null;
            for (iIndexofRowToRead = iIndexofStartRowToRead; iIndexofRowToRead <= shtTestDataSheet.getLastRowNum(); iIndexofRowToRead++) {
                row = shtTestDataSheet.getRow(iIndexofRowToRead);
                //System.out.println("Row Index: " +iIndexofRowToRead + ".. Value: " + row.getCell(0).getStringCellValue());
                
                try{
                	if (Cell.CELL_TYPE_STRING == row.getCell(iColIndex).getCellType()){
                		if (row.getCell(iColIndex).getStringCellValue().equalsIgnoreCase(strRowName)){
                            bRowNameFound = true;
                            iRowIndex = iIndexofRowToRead;
                            break;
                		}
                	}
                }
                catch(Exception e2){
                    System.out.println("In Catch for Row No: " +iIndexofRowToRead + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                }
                
            }
            
            if (!bRowNameFound)
                JOptionPane.showMessageDialog(null,"Row not found","Excel Error",JOptionPane.ERROR_MESSAGE); 
            
            //Close input stream
            inputStream.close();                
        }
        catch(Exception e){
                e.printStackTrace();
                JOptionPane.showMessageDialog(null,"Error in Excel operation","Excel Error",JOptionPane.ERROR_MESSAGE); 		
                iRowIndex = -1;
        }
        return iRowIndex;
    }  

}

