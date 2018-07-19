package Components;

import static GUI.PostDepTool.alALLProfilesinSFDC;
import static GUI.PostDepTool.alObjectProfilesfrmRTP;
import static GUI.PostDepTool.hmapProfileObjectAccessRecTypAssDefault;
import static Utilities.ExcelPOI.strTestDataFilePath;

import java.io.File;
import java.io.FileInputStream;
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

public class ReadObjectfromRTPExcel {
	public static int getStartRowIndexAfterAllProfTaginRTP_Object() throws Exception{
		int iRowAllProfileEnds = -1, iStartRow = -1;
		int iStartRowofObjectAccess = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Object Access:", 0, 0);
		iRowAllProfileEnds = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "ALL Profiles-End", 0, iStartRowofObjectAccess);
		if (iRowAllProfileEnds != -1)
			iStartRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Object-Start", 0, iRowAllProfileEnds) + 1;
		System.out.println(iRowAllProfileEnds);
		System.out.println(iStartRow);
		
		return iStartRow;
	}
	
	public static ArrayList<String> getListofObjectProfilesfrmRTP (int iStartRow) throws Exception{
		// get the list of Unique profiles from the Profile column in RTP
		ArrayList <String>alObjectProfilesfrmRTP = new ArrayList <String>();
		if (iStartRow != -1){
			int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Page Layout:", 0, iStartRow);
			
			System.out.println(iStartRow + ":" + iEndRow);
			alObjectProfilesfrmRTP = ExcelPOI.GetUniqueRowsfromColumn("VFPage,Class,Obj,PageLayout", "Profile", iStartRow, iEndRow);
			if (alObjectProfilesfrmRTP.contains("Profile"))
				alObjectProfilesfrmRTP.remove("Profile");
			if (alObjectProfilesfrmRTP.contains("Class Access:"))
				alObjectProfilesfrmRTP.remove("Class Access:");
			if (alObjectProfilesfrmRTP.contains("Object Access:"))
				alObjectProfilesfrmRTP.remove("Object Access:");
			if (alObjectProfilesfrmRTP.contains("Page Layout:"))
				alObjectProfilesfrmRTP.remove("Page Layout:");
			
			for (String str : alObjectProfilesfrmRTP){
				System.out.println(str);
			}
		}
		else
			FileOperations.writeToLog("All Profiles Start/End Tag not found");
		
		return alObjectProfilesfrmRTP;
	}

	//----- Get combination of the Field Obj Access for the Profiles in FLS Sheet (Except for the All profiles section) -----//
	public static void getCombinationofAllObjectProfinRTP (int iStartRow) throws Exception{
        String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "Object";
        String strColNameToRead_3 = "Access Level";
        String strColNameToRead_4 = "Record Type Assignment";
        String strColNameToRead_5 = "Default";
        
        int iRowToStartReadingFrom = iStartRow;
        int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Page Layout:", 0, iStartRow);
        String strSheetName = "VFPage,Class,Obj,PageLayout";
                      
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
            
            int n;
            boolean bColNameFound_5 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (n = 0; n <= row.getLastCellNum(); n++) {
                if (row.getCell(n).getStringCellValue().equals(strColNameToRead_5)){
                        bColNameFound_5 = true;
                        break;
                }
            }
            
            int i;
            if (bColNameFound_1 && bColNameFound_2 && bColNameFound_3 && bColNameFound_4 && bColNameFound_5){
            	hmapProfileObjectAccessRecTypAssDefault = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alObjectProfilesfrmRTP){
            		hmapProfileObjectAccessRecTypAssDefault.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alObject = new ArrayList <String> ();
                for (i = iStartRow + 1; i < iEndRow; i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueObject = row.getCell(0).getStringCellValue().trim();
                    			                    	
                    		if (!strCellValueObject.equals("")){
                    			// if First Col is Class-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueObject.equalsIgnoreCase("Object-End")){                    			
                        			alObject.add(strCellValueObject);
                    			}
                    			else{
                    				// Once the Field/Obj column is read and FieldObj Array is created for one FLS-Start FLS-End, 
                    				// Start reading the Profile column, pick up each profile and assign the FieldObj Array to each profile
                    				for (int iReadProfCol = iStartRow+1; iReadProfCol < i; iReadProfCol++){
                    					row = shtTestDataSheet.getRow(iReadProfCol);
                    					
                    					try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
		                    					String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
		                    					String strCellValueAccessLevel = "";
		                    					String strCellValueRecordTypeAssignment = "";
		                    					String strCellValueDefault = "";
		                    					try{ // Access level Column can be blank
			                    					strCellValueAccessLevel = row.getCell(l).getStringCellValue().trim();
		                    					}catch(Exception e){System.out.println("Access Level Col Catch");}
		                    					try{ // RecTypeAss Column can be blank
			                    					strCellValueRecordTypeAssignment = row.getCell(m).getStringCellValue().trim();
		                    					}catch(Exception e){System.out.println("RecTypeAss Level Col Catch");}
		                    					try{ // Default Column can be blank
			                    					strCellValueDefault = row.getCell(n).getStringCellValue().trim();
		                    					}catch(Exception e){System.out.println("Default Col Catch");}
		                    					
		                    					if (strCellValueAccessLevel.equals(""))
		                    						strCellValueAccessLevel = "BLANK";
		                    					if (strCellValueRecordTypeAssignment.equals(""))
		                    						strCellValueRecordTypeAssignment = "BLANK";
		                    					if (strCellValueDefault.equals(""))
		                    						strCellValueDefault = "BLANK";
		                    					
		                    					if (!strCellValueObject.equals("")){
			                    					// If a profile is already allotted an array of Field/Obj from a prev FLS Start FLS End, then 
			                    					// append the current array of Field/Obj to the Profile's hmap, else add the new Array
			                    					// Also include the AccessLevel 
			                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
			                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
			                    					alAppendArrayToExistingEntries = hmapProfileObjectAccessRecTypAssDefault.get(strCellValueProfile);
			                    					
			                    					if (alAppendArrayToExistingEntries.size() > 0){
			                    						for (String str : alAppendArrayToExistingEntries){
			                    							//str = str + "#" + strCellValueAccess;
			                    							alfinalArray.add(str);
			                    						}
			                    					}
			                    					for (String str : alObject){
			                    						str = str + "#" + strCellValueAccessLevel + "#" + strCellValueRecordTypeAssignment + "#" + strCellValueDefault;
			                							alfinalArray.add(str);
			                						}
			                    					hmapProfileObjectAccessRecTypAssDefault.put(strCellValueProfile, alfinalArray);
		                    					}
                                        	}
                                        }
                                        catch (Exception e){
                                        	System.out.println("Possible blank rows before Object-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT Class-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        	FileOperations.writeToLog("Possible blank rows before Object-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT Class-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        }
                    				}
                    				
                    				// Reset the field/obj array
                    				// Determine the next Object-Start row and set the i Counter and StartRow counter
                    				alObject = new ArrayList <String> ();
                    				for (int iNextFLSStartRow = i; iNextFLSStartRow < shtTestDataSheet.getLastRowNum(); iNextFLSStartRow++){
                    					row = shtTestDataSheet.getRow(iNextFLSStartRow);
                                    	try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                                        		strCellValueObject = row.getCell(0).getStringCellValue().trim();
                                        		if (strCellValueObject.equalsIgnoreCase("Object-Start")){
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
                
                Set set = hmapProfileObjectAccessRecTypAssDefault.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("key is: "+ mentry.getKey() + " & Value is: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("key is: "+ mentry.getKey() + " & Value is: ");
        			FileOperations.writeToLog(mentry.getValue().toString());
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

	//----- Get combination of the Field Obj Access for the All profiles section in VFPage,Class,Obj,PageLayout Sheet -----//
	public static void getCombinationofObjectfrmALLPROFILES (int iStartRowAllProf) throws Exception { 
        String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "Object";
        String strColNameToRead_3 = "Access Level";
        String strColNameToRead_4 = "Record Type Assignment";
        String strColNameToRead_5 = "Default";
        
        int iRowToStartReadingFrom = iStartRowAllProf;
        int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Page Layout:", 0, iStartRowAllProf)-1;
        String strSheetName = "VFPage,Class,Obj,PageLayout";
                      
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
            
            int n;
            boolean bColNameFound_5 = false;
            row = shtTestDataSheet.getRow(iRowToStartReadingFrom);
            for (n = 0; n <= row.getLastCellNum(); n++) {
                if (row.getCell(n).getStringCellValue().equals(strColNameToRead_5)){
                        bColNameFound_5 = true;
                        break;
                }
            }

            int i;
            if (bColNameFound_1 && bColNameFound_2 && bColNameFound_3 && bColNameFound_4 && bColNameFound_5){
            	//hmapProfileVFPageAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alALLProfilesinSFDC){
            		if (!hmapProfileObjectAccessRecTypAssDefault.containsKey(strProfile))
            			hmapProfileObjectAccessRecTypAssDefault.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alObject = new ArrayList <String> ();
                for (i = iStartRowAllProf + 1; i < iEndRow; i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueObject = row.getCell(0).getStringCellValue().trim();
        					String strCellValueAccessLevel = "";
        					String strCellValueRecordTypeAssignment = "";
        					String strCellValueDefault = "";
                    			                    	
                    		if (!strCellValueObject.equals("")){
                    			// if First Col is VFPage-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueObject.equalsIgnoreCase("ALL Profiles-End")){  
                					try{ // Access level Column can be blank
                    					strCellValueAccessLevel = row.getCell(l).getStringCellValue().trim();
                					}catch(Exception e){System.out.println("Access Level Col Catch");}
                					try{ // RecTypeAss Column can be blank
                    					strCellValueRecordTypeAssignment = row.getCell(m).getStringCellValue().trim();
                					}catch(Exception e){System.out.println("RecTypeAss Level Col Catch");}
                					try{ // Default Column can be blank
                    					strCellValueDefault = row.getCell(n).getStringCellValue().trim();
                					}catch(Exception e){System.out.println("Default Col Catch");}
                					
                					if (strCellValueAccessLevel.equals(""))
                						strCellValueAccessLevel = "BLANK";
                					if (strCellValueRecordTypeAssignment.equals(""))
                						strCellValueRecordTypeAssignment = "BLANK";
                					if (strCellValueDefault.equals(""))
                						strCellValueDefault = "BLANK";
                					
                					String strAccessLvlRecTypeDef = strCellValueObject + "#" + strCellValueAccessLevel + "#" + strCellValueRecordTypeAssignment + "#" + strCellValueDefault;
                					
                    				for (String strProfile : alALLProfilesinSFDC){
                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
                    					alAppendArrayToExistingEntries = hmapProfileObjectAccessRecTypAssDefault.get(strProfile);
                    					
                    					if (alAppendArrayToExistingEntries.size() > 0){
                    						for (String str : alAppendArrayToExistingEntries){
                    							//str = str + "#" + strCellValueAccess;
                    							alfinalArray.add(str);
                    						}
                    					}
                    					alfinalArray.add(strAccessLvlRecTypeDef);                						
                    					hmapProfileObjectAccessRecTypAssDefault.put(strProfile, alfinalArray); 
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
                
        		hmapProfileObjectAccessRecTypAssDefault.put( "Standard", hmapProfileObjectAccessRecTypAssDefault.remove( "Standard User" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "Admin", hmapProfileObjectAccessRecTypAssDefault.remove( "System Administrator" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "ContractManager", hmapProfileObjectAccessRecTypAssDefault.remove( "Contract Manager" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "MarketingProfile", hmapProfileObjectAccessRecTypAssDefault.remove( "Marketing User" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "ReadOnly", hmapProfileObjectAccessRecTypAssDefault.remove( "Read Only" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "SolutionManager", hmapProfileObjectAccessRecTypAssDefault.remove( "Solution Manager" ) );
        		hmapProfileObjectAccessRecTypAssDefault.put( "StandardAul", hmapProfileObjectAccessRecTypAssDefault.remove( "Standard Platform User" ) );

                Set set = hmapProfileObjectAccessRecTypAssDefault.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("PROFILE: "+ mentry.getKey() + ".... COMBINATIONS: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("-------------------------------------------------");
        			FileOperations.writeToLog("Profile: "+ mentry.getKey());
        			FileOperations.writeToLog("Object: "+ mentry.getValue().toString());
        		}
            }
            else{
            	//System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","Profile/Class Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
	}

}
