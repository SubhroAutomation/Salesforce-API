package Components;

import static GUI.PostDepTool.alALLProfilesinSFDC;
import static GUI.PostDepTool.alFLSProfilesfrmRTP;
import static GUI.PostDepTool.hmapProfileFieldObjAccess;
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

public class ReadFLSfromRTPExcel {
	//------------ Get starting row to read frm RTP Excel (After All Profiles) ------------//
	public static int getStartRowIndexAfterAllProfTaginRTP() throws Exception{
		int iRowAllProfileEnds = -1, iStartRow = -1;
		iRowAllProfileEnds = ExcelPOI.GetRowIndexofValueinCol("FLS", "ALL Profiles-End", 0, 0);
		if (iRowAllProfileEnds != -1)
			iStartRow = ExcelPOI.GetRowIndexofValueinCol("FLS", "FLS-Start", 0, iRowAllProfileEnds) + 1;
		System.out.println(iRowAllProfileEnds);
		System.out.println(iStartRow);
		
		return iStartRow;
	}
	
	//--------------------- Get list of All Profiles from the FLS Sheet -------------------//
	public static ArrayList<String> getListofFLSProfilesfrmRTP (int iStartRow) throws IOException{
		// get the list of Unique profiles from the Profile column in RTP
		ArrayList <String>alFLSProfilesfrmRTP = new ArrayList <String>();
		if (iStartRow != -1){
			alFLSProfilesfrmRTP = ExcelPOI.GetUniqueRowsfromColumn("FLS", "Profile", iStartRow);
			if (alFLSProfilesfrmRTP.contains("Profile"))
				alFLSProfilesfrmRTP.remove("Profile");
			
			for (String str : alFLSProfilesfrmRTP){
				//System.out.println(str);
			}
		}
		else
			FileOperations.writeToLog("All Profiles Start/End Tag not found");
		
		return alFLSProfilesfrmRTP;
	}
	
	//----- Get combination of the Field Obj Access for the Profiles in FLS Sheet (Except for the All profiles section) -----//
	public static void getCombinationofAllFieldObjAccessProfinRTP (int iStartRow) throws IOException{
        String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "API Name";
        String strColNameToRead_3 = "Parent Object";
        String strColNameToRead_4 = "Access Level";
        int iRowToStartReadingFrom = iStartRow;
        String strSheetName = "FLS";
                      
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
            	hmapProfileFieldObjAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alFLSProfilesfrmRTP){
            		hmapProfileFieldObjAccess.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
                for (i = iStartRow + 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueField = row.getCell(0).getStringCellValue().trim();
                    		String strCellValueFieldAPI = "";
	                    	String strCellValueObject = "";
                    			                    	
                    		if (!strCellValueField.equals("")){
                    			// if First Col is FLS-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueField.equalsIgnoreCase("FLS-End")){                    			
    		                    	strCellValueFieldAPI = row.getCell(k).getStringCellValue().trim();
    		                    	strCellValueObject = row.getCell(l).getStringCellValue().trim();
    	                    	
                    				String strFieldObjAccess = strCellValueObject + "." + strCellValueFieldAPI; // + "#" + strCellValueAccess;
	                    			alFieldObjectAccess.add(strFieldObjAccess);
                    			}
                    			else{
                    				// Once the Field/Obj column is read and FieldObj Array is created for one FLS-Start FLS-End, 
                    				// Start reading the Profile column, pick up each profile and assign the FieldObj Array to each profile
                    				for (int iReadProfCol = iStartRow+1; iReadProfCol < i; iReadProfCol++){
                    					row = shtTestDataSheet.getRow(iReadProfCol);
                    					
                    					try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
		                    					String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
		                    					String strCellValueAccess = row.getCell(m).getStringCellValue().trim();
		                    					  
		                    					if (!strCellValueField.equals("")){
			                    					// If a profile is already allotted an array of Field/Obj from a prev FLS Start FLS End, then 
			                    					// append the current array of Field/Obj to the Profile's hmap, else add the new Array
			                    					// Also include the AccessLevel 
			                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
			                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
			                    					alAppendArrayToExistingEntries = hmapProfileFieldObjAccess.get(strCellValueProfile);
			                    					
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
			                    					hmapProfileFieldObjAccess.put(strCellValueProfile, alfinalArray);
		                    					}
                                        	}
                                        }
                                        catch (Exception e){
                                        	System.out.println("Possible blank rows before FLS-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT FLS-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        	FileOperations.writeToLog("Possible blank rows before FLS-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT FLS-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        }
                    				}
                    				
                    				// Reset the field/obj array
                    				// Determine the next FLS-Start row and set the i Counter and StartRow counter
                    				alFieldObjectAccess = new ArrayList <String> ();
                    				for (int iNextFLSStartRow = i; iNextFLSStartRow < shtTestDataSheet.getLastRowNum(); iNextFLSStartRow++){
                    					row = shtTestDataSheet.getRow(iNextFLSStartRow);
                                    	try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                                        		strCellValueField = row.getCell(0).getStringCellValue().trim();
                                        		if (strCellValueField.equalsIgnoreCase("FLS-Start")){
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
                
                /*Set set = hmapProfileFieldObjAccess.entrySet();
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
                JOptionPane.showMessageDialog(null,"Column Missing!","Profile/API Name/Object/Access level Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
    }

	//----- Get combination of the Field Obj Access for the All profiles section in FLS Sheet -----//
	public static void getCombinationofFieldObjAccfrmALLPROFILES (int iStartRowAllProf) throws IOException { 
		String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "API Name";
        String strColNameToRead_3 = "Parent Object";
        String strColNameToRead_4 = "Access Level";
        int iRowToStartReadingFrom = iStartRowAllProf;
        String strSheetName = "FLS";
                      
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
            	//hmapProfileFieldObjAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alALLProfilesinSFDC){
            		if (!hmapProfileFieldObjAccess.containsKey(strProfile))
            			hmapProfileFieldObjAccess.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
                for (i = iStartRowAllProf + 1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueField = row.getCell(0).getStringCellValue().trim();
                    		String strCellValueFieldAPI = "";
	                    	String strCellValueObject = "";
	                    	String strCellValueAccess = "";
                    			                    	
                    		if (!strCellValueField.equals("")){
                    			// if First Col is FLS-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueField.equalsIgnoreCase("ALL Profiles-End")){                    			
    		                    	strCellValueFieldAPI = row.getCell(k).getStringCellValue().trim();
    		                    	strCellValueObject = row.getCell(l).getStringCellValue().trim();
    		                    	strCellValueAccess = row.getCell(m).getStringCellValue().trim();
    	                    	
                    				String strFieldObjAccess = strCellValueObject + "." + strCellValueFieldAPI + "#" + strCellValueAccess;

                    				for (String strProfile : alALLProfilesinSFDC){
                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
                    					alAppendArrayToExistingEntries = hmapProfileFieldObjAccess.get(strProfile);
                    					
                    					if (alAppendArrayToExistingEntries.size() > 0){
                    						for (String str : alAppendArrayToExistingEntries){
                    							//str = str + "#" + strCellValueAccess;
                    							alfinalArray.add(str);
                    						}
                    					}
                    					alfinalArray.add(strFieldObjAccess);                						
                    					hmapProfileFieldObjAccess.put(strProfile, alfinalArray); 
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
                
                hmapProfileFieldObjAccess.put( "Standard", hmapProfileFieldObjAccess.remove( "Standard User" ) );
        		hmapProfileFieldObjAccess.put( "Admin", hmapProfileFieldObjAccess.remove( "System Administrator" ) );
        		hmapProfileFieldObjAccess.put( "ContractManager", hmapProfileFieldObjAccess.remove( "Contract Manager" ) );
        		hmapProfileFieldObjAccess.put( "MarketingProfile", hmapProfileFieldObjAccess.remove( "Marketing User" ) );
        		hmapProfileFieldObjAccess.put( "ReadOnly", hmapProfileFieldObjAccess.remove( "Read Only" ) );
        		hmapProfileFieldObjAccess.put( "SolutionManager", hmapProfileFieldObjAccess.remove( "Solution Manager" ) );
        		hmapProfileFieldObjAccess.put( "StandardAul", hmapProfileFieldObjAccess.remove( "Standard Platform User" ) );
        		
                Set set = hmapProfileFieldObjAccess.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("PROFILE: "+ mentry.getKey() + ".... COMBINATIONS: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("-------------------------------------------------");
        			FileOperations.writeToLog("Profile: "+ mentry.getKey());
        			FileOperations.writeToLog("Field/Obj/Access Entries: "+ mentry.getValue().toString());
        		}
            }
            else{
            	//System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","Profile/API Name/Object/Access level Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
	}

	//----- Not calling this function since it will take a long time to create the map entries in excel -----//
	public static void createProfileFieldObjAccessSheet () throws IOException {
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
            
            
            Set set = hmapProfileFieldObjAccess.entrySet();
    		Iterator iterator = set.iterator();
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			FileOperations.writeToLog("Profile: "+ mentry.getKey());
    			FileOperations.writeToLog("Field/Obj/Access Entries: "+ mentry.getValue().toString());
    			
    			ArrayList <String> alFieldObjectAccess = new ArrayList <String> ();
    			alFieldObjectAccess = hmapProfileFieldObjAccess.get(mentry.getKey());
    			
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
