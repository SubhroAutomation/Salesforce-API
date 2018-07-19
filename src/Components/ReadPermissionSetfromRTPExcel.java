package Components;

import static GUI.PostDepTool.alALLPermissionSetsinSFDC;
import static GUI.PostDepTool.alPermissionSetsfrmRTP;
import static GUI.PostDepTool.hmapPermissionSetFieldObjAccess;
import static Utilities.ExcelPOI.strTestDataFilePath;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Utilities.ExcelPOI;
import Utilities.FileOperations;

public class ReadPermissionSetfromRTPExcel {
	//------------ Get starting row to read frm RTP Excel (After All Permission Sets) ------------//
	public static int getStartRowIndexAfterAllPermissionTaginRTP() throws Exception{
		int iRowAllPermissionSetEnds = -1, iStartRow = -1;
		iRowAllPermissionSetEnds = ExcelPOI.GetRowIndexofValueinCol("Permission Set changes", "ALL Permission Set-End", 0, 0);
		if (iRowAllPermissionSetEnds != -1)
			iStartRow = ExcelPOI.GetRowIndexofValueinCol("Permission Set changes", "PermissionSet-Start", 0, iRowAllPermissionSetEnds) + 1;
		System.out.println(iRowAllPermissionSetEnds);
		System.out.println(iStartRow);
		
		return iStartRow;
	}
	
	//--------------------- Get list of All PermissionSets from the FLS Sheet -------------------//
	public static ArrayList<String> getListofPermissionSetsfrmRTP (int iStartRow) throws IOException{
		// get the list of Unique PermissionSets from the PermissionSet column in RTP
		ArrayList <String>alPermissionSetsfrmRTP = new ArrayList <String>();
		if (iStartRow != -1){
			alPermissionSetsfrmRTP = ExcelPOI.GetUniqueRowsfromColumn("Permission Set changes", "Permission Set", iStartRow);
			if (alPermissionSetsfrmRTP.contains("Permission Set"))
				alPermissionSetsfrmRTP.remove("Permission Set");
			
			for (String str : alPermissionSetsfrmRTP){
				//System.out.println(str);
			}
		}
		else
			FileOperations.writeToLog("All Permission Set Start/End Tag not found");
		
		return alPermissionSetsfrmRTP;
	}
	
	//----- Get combination of the Field Obj Access for the PermissionSets in FLS Sheet (Except for the All PermissionSets section) -----//
	public static void getCombinationofAllFieldObjAccessPermSetinRTP (int iStartRow) throws IOException{
        String strColNameToRead_1 = "Permission Set";
        String strColNameToRead_2 = "API Name";
        String strColNameToRead_3 = "Object";
        String strColNameToRead_4 = "Access Level";
        int iRowToStartReadingFrom = iStartRow;
        String strSheetName = "Permission Set changes";
                      
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

            int j;
            boolean bColNameFound_1 = false;
            Row row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (j = 0; j <= row.getLastCellNum(); j++) {
                if (row.getCell(j).getStringCellValue().equals(strColNameToRead_1)){
                        bColNameFound_1 = true;
                        break;
                }
            }
            
            int k;
            boolean bColNameFound_2 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (k = 0; k <= row.getLastCellNum(); k++) {
                if (row.getCell(k).getStringCellValue().equals(strColNameToRead_2)){
                        bColNameFound_2 = true;
                        break;
                }
            }

            int l;
            boolean bColNameFound_3 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (l = 0; l <= row.getLastCellNum(); l++) {
                if (row.getCell(l).getStringCellValue().equals(strColNameToRead_3)){
                        bColNameFound_3 = true;
                        break;
                }
            }
            
            int m;
            boolean bColNameFound_4 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (m = 0; m <= row.getLastCellNum(); m++) {
                if (row.getCell(m).getStringCellValue().equals(strColNameToRead_4)){
                        bColNameFound_4 = true;
                        break;
                }
            }

            int i;
            if (bColNameFound_1 && bColNameFound_2 && bColNameFound_3 && bColNameFound_4){
            	hmapPermissionSetFieldObjAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique PermissionSets from RTP
            	for (String strPermissionSet : alPermissionSetsfrmRTP){
            		hmapPermissionSetFieldObjAccess.put(strPermissionSet, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
                for (i = iStartRow + 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValuePermissionSet = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueFieldAPI = row.getCell(0).getStringCellValue().trim();
	                    	String strCellValueObject = "";
                    			                    	
                    		if (!strCellValueFieldAPI.equals("")){
                    			// if First Col is FLS-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueFieldAPI.equalsIgnoreCase("PermissionSet-End")){                    			
    		                    	strCellValueObject = row.getCell(l).getStringCellValue().trim();
    	                    	
                    				String strFieldObjAccess = strCellValueObject + "." + strCellValueFieldAPI; // + "#" + strCellValueAccess;
	                    			alFieldObjectAccess.add(strFieldObjAccess);
                    			}
                    			else{
                    				// Once the Field/Obj column is read and FieldObj Array is created for one FLS-Start FLS-End, 
                    				// Start reading the PermissionSet column, pick up each PermissionSet and assign the FieldObj Array to each PermissionSet
                    				for (int iReadProfCol = iStartRow+1; iReadProfCol < i; iReadProfCol++){
                    					row = shtTestDataSheet.getRow(iReadProfCol);
                    					
                    					try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
		                    					String strCellValuePermissionSet = row.getCell(j).getStringCellValue().trim();
		                    					String strCellValueAccess = row.getCell(m).getStringCellValue().trim();
		                    					  
		                    					if (!strCellValueFieldAPI.equals("")){
			                    					// If a PermissionSet is already allotted an array of Field/Obj from a prev FLS Start FLS End, then 
			                    					// append the current array of Field/Obj to the PermissionSet's hmap, else add the new Array
			                    					// Also include the AccessLevel 
			                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
			                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
			                    					alAppendArrayToExistingEntries = hmapPermissionSetFieldObjAccess.get(strCellValuePermissionSet);
			                    					
			                    					if (alAppendArrayToExistingEntries.size() > 0){
			                    						for (String str : alAppendArrayToExistingEntries){
			                    							//str = str + "#" + strCellValueAccess;
			                    							alfinalArray.add(str);
			                    						}
			                    					}
			                						for (String str : alFieldObjectAccess){
			                							str = str + "#" + strCellValueAccess;
			                							alfinalArray.add(str);
			                						}
			                						hmapPermissionSetFieldObjAccess.put(strCellValuePermissionSet, alfinalArray);
		                    					}
                                        	}
                                        }
                                        catch (Exception e){
                                        	System.out.println("Possible blank rows before PermissionSet-END, i.e, after Field/Obj/PermissionSet entries ended, next immediate row is NOT PermissionSet-END, hence PermissionSet column will contain blank values in the end.. row: " +iReadProfCol);
                                        	FileOperations.writeToLog("Possible blank rows before PermissionSet-END, i.e, after Field/Obj/PermissionSet entries ended, next immediate row is NOT PermissionSet-END, hence PermissionSet column will contain blank values in the end.. row: " +iReadProfCol);
                                        }
                    				}
                    				
                    				// Reset the field/obj array
                    				// Determine the next FLS-Start row and set the i Counter and StartRow counter
                    				alFieldObjectAccess = new ArrayList <String> ();
                    				for (int iNextFLSStartRow = i; iNextFLSStartRow < shtTestDataSheet.getLastRowNum(); iNextFLSStartRow++){
                    					row = shtTestDataSheet.getRow(iNextFLSStartRow);
                                    	try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                                        		strCellValueFieldAPI = row.getCell(0).getStringCellValue().trim();
                                        		if (strCellValueFieldAPI.equalsIgnoreCase("PermissionSet-Start")){
                                        			//int iFLSStartRow = iNextFLSStartRow;
                                        			iStartRow = iNextFLSStartRow + 1;
                                        			i = iNextFLSStartRow + 1;
                                        			break;
                                        		}
                                        	}
                                    	} catch (Exception e){
                                    		System.out.println("Inside iNextFLSStartRow: Blank Cell possible reason.. Row: " +iNextFLSStartRow);
                                    		FileOperations.writeToLog("Inside iNextFLSStartRow: Blank Cell possible reason.. Row: " +iNextFLSStartRow);
                                    	}
                    				}
                    			}
                    		}
                    	}
                    }
                    catch(Exception e2){
                        System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                        FileOperations.writeToLog("Inside Catch: " +e2.getMessage());
                    }
                }
                
                /*Set set = hmapPermissionSetFieldObjAccess.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("key is: "+ mentry.getKey() + " & Value is: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("key is: "+ mentry.getKey() + " & Value is: ");
        			FileOperations.writeToLog(mentry.getValue().toString());
        		}*/
            }
            else{
            	//System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","PermissionSet/API Name/Object/Access level Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
    }

	//----- Get combination of the Field Obj Access for the All PermissionSets section in FLS Sheet -----//
	public static void getCombinationofFieldObjAccfrmALLPERMISSIONSETS (int iStartRowAllProf) throws IOException { 
		String strColNameToRead_1 = "Permission Set";
        String strColNameToRead_2 = "API Name";
        String strColNameToRead_3 = "Object";
        String strColNameToRead_4 = "Access Level";
        int iRowToStartReadingFrom = iStartRowAllProf;
        String strSheetName = "Permission Set changes";
                      
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

            int j;
            boolean bColNameFound_1 = false;
            Row row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (j = 0; j <= row.getLastCellNum(); j++) {
                if (row.getCell(j).getStringCellValue().equals(strColNameToRead_1)){
                        bColNameFound_1 = true;
                        break;
                }
            }
            
            int k;
            boolean bColNameFound_2 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (k = 0; k <= row.getLastCellNum(); k++) {
                if (row.getCell(k).getStringCellValue().equals(strColNameToRead_2)){
                        bColNameFound_2 = true;
                        break;
                }
            }

            int l;
            boolean bColNameFound_3 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (l = 0; l <= row.getLastCellNum(); l++) {
                if (row.getCell(l).getStringCellValue().equals(strColNameToRead_3)){
                        bColNameFound_3 = true;
                        break;
                }
            }
            
            int m;
            boolean bColNameFound_4 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (m = 0; m <= row.getLastCellNum(); m++) {
                if (row.getCell(m).getStringCellValue().equals(strColNameToRead_4)){
                        bColNameFound_4 = true;
                        break;
                }
            }

            int i;
            if (bColNameFound_1 && bColNameFound_2 && bColNameFound_3 && bColNameFound_4){
            	//hmapPermissionSetFieldObjAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique PermissionSets from RTP
            	for (String strPermissionSet : alALLPermissionSetsinSFDC){
            		if (!hmapPermissionSetFieldObjAccess.containsKey(strPermissionSet))
            			hmapPermissionSetFieldObjAccess.put(strPermissionSet, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
                for (i = iStartRowAllProf + 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValuePermissionSet = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueFieldAPI = row.getCell(0).getStringCellValue().trim();
	                    	String strCellValueObject = "";
	                    	String strCellValueAccess = "";
                    			                    	
                    		if (!strCellValueFieldAPI.equals("")){
                    			// if First Col is FLS-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueFieldAPI.equalsIgnoreCase("ALL Permission Set-End")){                    			
    		                    	strCellValueObject = row.getCell(l).getStringCellValue().trim();
    		                    	strCellValueAccess = row.getCell(m).getStringCellValue().trim();
    	                    	
                    				String strFieldObjAccess = strCellValueObject + "." + strCellValueFieldAPI + "#" + strCellValueAccess;

                    				for (String strPermissionSet : alALLPermissionSetsinSFDC){
                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
                    					alAppendArrayToExistingEntries = hmapPermissionSetFieldObjAccess.get(strPermissionSet);
                    					
                    					if (alAppendArrayToExistingEntries.size() > 0){
                    						for (String str : alAppendArrayToExistingEntries){
                    							//str = str + "#" + strCellValueAccess;
                    							alfinalArray.add(str);
                    						}
                    					}
                    					alfinalArray.add(strFieldObjAccess);                						
                    					hmapPermissionSetFieldObjAccess.put(strPermissionSet, alfinalArray); 
                    				}
                    			}
                    			else{
                    				break;
                    			}
                    		}
                    	}
                    }
                    catch(Exception e2){
                        System.out.println("In Catch for Row No: " +i + ".. Possible Reason: Value in Cell is Not CELL_TYPE_STRING");
                        FileOperations.writeToLog("Inside Catch: " +e2.getMessage());
                    }
                }
                
                Set set = hmapPermissionSetFieldObjAccess.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("PERMISSION SET: "+ mentry.getKey() + ".... COMBINATIONS: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("-------------------------------------------------");
        			FileOperations.writeToLog("Permission Set: "+ mentry.getKey());
        			FileOperations.writeToLog("Field/Obj/Access Entries: "+ mentry.getValue().toString());
        		}
            }
            else{
            	//System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","PermissionSet/API Name/Object/Access level Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
	}

	//----- Not calling this function since it will take a long time to create the map entries in excel -----//
	public static void createPermissionSetFieldObjAccessSheet () throws IOException {
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

            //Create new sheet
            String strSheetProfilesbasedFLS = "ProfilesbasedFLS";
            ExcelPOI.AddNewSheet(strSheetProfilesbasedFLS);
            Sheet shtTestDataSheet = wbTestDataExcelWB.getSheet(strSheetProfilesbasedFLS);
            
            int iRow = 0;
            shtTestDataSheet.createRow(iRow);
            
            Row row = shtTestDataSheet.getRow(iRow);
            Cell cell = row.createCell(0);
            cell.setCellValue("Profile");
            cell = row.createCell(1);
            cell.setCellValue("Field API");
            cell = row.createCell(2);
            cell.setCellValue("Object");
            cell = row.createCell(3);
            cell.setCellValue("Access Level");
            cell = row.createCell(4);
            cell.setCellValue("Status");
            
            //Close input stream
            inputStream.close();

            //Create an object of FileOutputStream class to create write data in excel file
            FileOutputStream outputStream = new FileOutputStream(file);

            //write data in the excel file
            wbTestDataExcelWB.write(outputStream);

            //close output stream
            outputStream.close();
            
            
            Set set = hmapPermissionSetFieldObjAccess.entrySet();
    		Iterator iterator = set.iterator();
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			FileOperations.writeToLog("Profile: "+ mentry.getKey());
    			FileOperations.writeToLog("Field/Obj/Access Entries: "+ mentry.getValue().toString());
    			
    			ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
    			alFieldObjectAccess = hmapPermissionSetFieldObjAccess.get(mentry.getKey());
    			
    			System.out.println("Profile: " +mentry.getKey().toString() + "...................");
    			for (String strFOA : alFieldObjectAccess){
    				strFOA = strFOA.replace(".", "#");
    				String strObject = strFOA.split("#")[0];
    				String strField = strFOA.split("#")[1];
    				String strAccess = strFOA.split("#")[2];
    				
    				iRow ++;        			
    				ExcelPOI.WriteDataToExcel(strSheetProfilesbasedFLS, "Profile", iRow, mentry.getKey().toString());
    				ExcelPOI.WriteDataToExcel(strSheetProfilesbasedFLS, "Field API", iRow, strField);
    				ExcelPOI.WriteDataToExcel(strSheetProfilesbasedFLS, "Object", iRow, strObject);
    				ExcelPOI.WriteDataToExcel(strSheetProfilesbasedFLS, "Access Level", iRow, strAccess);    				
    			}
    		}
		}
		catch(Exception e){
            e.printStackTrace();
            FileOperations.writeToLog("Inside Catch : createProfileFieldObjAccessSheet");
		}
	}

	
}
