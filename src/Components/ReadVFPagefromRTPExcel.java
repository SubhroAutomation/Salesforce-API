package Components;

import static GUI.PostDepTool.alALLProfilesinSFDC;
import static GUI.PostDepTool.alVFPageProfilesfrmRTP;
import static GUI.PostDepTool.hmapProfileVFPageAccess;
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

public class ReadVFPagefromRTPExcel {
	//------------ Get starting row to read frm RTP Excel (After All Profiles) ------------//
	public static int getStartRowIndexAfterAllProfTaginRTP_VFPage() throws Exception{
		int iRowAllProfileEnds = -1, iStartRow = -1;
		iRowAllProfileEnds = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "ALL Profiles-End", 0, 0);
		if (iRowAllProfileEnds != -1)
			iStartRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "VFPage-Start", 0, iRowAllProfileEnds) + 1;
		System.out.println(iRowAllProfileEnds);
		System.out.println(iStartRow);
		
		return iStartRow;
	}

	//--------------------- Get list of All Profiles from the FLS Sheet -------------------//
	public static ArrayList<String> getListofVFPageProfilesfrmRTP (int iStartRow) throws Exception{
		// get the list of Unique profiles from the Profile column in RTP
		ArrayList <String>alVFPageProfilesfrmRTP = new ArrayList <String>();
		if (iStartRow != -1){
			int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Class Access:", 0, iStartRow);
			
			System.out.println(iStartRow + ":" + iEndRow);
			alVFPageProfilesfrmRTP = ExcelPOI.GetUniqueRowsfromColumn("VFPage,Class,Obj,PageLayout", "Profile", iStartRow, iEndRow);
			if (alVFPageProfilesfrmRTP.contains("Profile"))
				alVFPageProfilesfrmRTP.remove("Profile");
			if (alVFPageProfilesfrmRTP.contains("Visualforce Page Access:"))
				alVFPageProfilesfrmRTP.remove("Visualforce Page Access:");
			if (alVFPageProfilesfrmRTP.contains("Class Access:"))
				alVFPageProfilesfrmRTP.remove("Class Access:");
			
			for (String str : alVFPageProfilesfrmRTP){
				System.out.println(str);
			}
		}
		else
			FileOperations.writeToLog("All Profiles Start/End Tag not found");
		
		return alVFPageProfilesfrmRTP;
	}

	//----- Get combination of the Field Obj Access for the Profiles in FLS Sheet (Except for the All profiles section) -----//
	public static void getCombinationofAllVFPageProfinRTP (int iStartRow) throws Exception{
        String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "Page";
        int iRowToStartReadingFrom = iStartRow;
        int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Class Access:", 0, iStartRow);
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

            int i;
            if (bColNameFound_1 && bColNameFound_2){
            	hmapProfileVFPageAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alVFPageProfilesfrmRTP){
            		hmapProfileVFPageAccess.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alVFPage = new ArrayList <String> ();
                for (i = iStartRow + 1; i < iEndRow; i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueVFPage = row.getCell(0).getStringCellValue().trim();
                    			                    	
                    		if (!strCellValueVFPage.equals("")){
                    			// if First Col is FLS-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueVFPage.equalsIgnoreCase("VFPage-End")){                    			
                        			alVFPage.add(strCellValueVFPage);
                    			}
                    			else{
                    				// Once the Field/Obj column is read and FieldObj Array is created for one FLS-Start FLS-End, 
                    				// Start reading the Profile column, pick up each profile and assign the FieldObj Array to each profile
                    				for (int iReadProfCol = iStartRow+1; iReadProfCol < i; iReadProfCol++){
                    					row = shtTestDataSheet.getRow(iReadProfCol);
                    					
                    					try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
		                    					String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
		                    					  
		                    					if (!strCellValueVFPage.equals("")){
			                    					// If a profile is already allotted an array of Field/Obj from a prev FLS Start FLS End, then 
			                    					// append the current array of Field/Obj to the Profile's hmap, else add the new Array
			                    					// Also include the AccessLevel 
			                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
			                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
			                    					alAppendArrayToExistingEntries = hmapProfileVFPageAccess.get(strCellValueProfile);
			                    					
			                    					if (alAppendArrayToExistingEntries.size() > 0){
			                    						for (String str : alAppendArrayToExistingEntries){
			                    							//str = str + "#" + strCellValueAccess;
			                    							alfinalArray.add(str);
			                    						}
			                    					}
			                    					for (String str : alVFPage){
			                							alfinalArray.add(str);
			                						}
			                    					hmapProfileVFPageAccess.put(strCellValueProfile, alfinalArray);
		                    					}
                                        	}
                                        }
                                        catch (Exception e){
                                        	System.out.println("Possible blank rows before VFPage-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT VFPage-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        	FileOperations.writeToLog("Possible blank rows before VFPage-END, i.e, after Field/Obj/Profile entries ended, next immediate row is NOT VFPage-END, hence Profile column will contain blank values in the end.. row: " +iReadProfCol);
                                        }
                    				}
                    				
                    				// Reset the field/obj array
                    				// Determine the next VFPage-Start row and set the i Counter and StartRow counter
                    				alVFPage = new ArrayList <String> ();
                    				for (int iNextFLSStartRow = i; iNextFLSStartRow < shtTestDataSheet.getLastRowNum(); iNextFLSStartRow++){
                    					row = shtTestDataSheet.getRow(iNextFLSStartRow);
                                    	try{
                                        	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                                        		strCellValueVFPage = row.getCell(0).getStringCellValue().trim();
                                        		if (strCellValueVFPage.equalsIgnoreCase("VFPage-Start")){
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
                
                Set set = hmapProfileVFPageAccess.entrySet();
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
	public static void getCombinationofVFPagefrmALLPROFILES (int iStartRowAllProf) throws Exception { 
		String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "Page";
        int iRowToStartReadingFrom = iStartRowAllProf;
        int iEndRow = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Class Access:", 0, iStartRowAllProf)-1;
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

            int i;
            if (bColNameFound_1 && bColNameFound_2){
            	//hmapProfileVFPageAccess = new LinkedHashMap<> ();
            	// Populate the map with all unique profiles from RTP
            	for (String strProfile : alALLProfilesinSFDC){
            		if (!hmapProfileVFPageAccess.containsKey(strProfile))
            			hmapProfileVFPageAccess.put(strProfile, new ArrayList<String>());
            	}
            	
            	ArrayList <String> alVFPage = new ArrayList <String> ();
                for (i = iStartRowAllProf + 1; i < iEndRow; i++) {
                	row = shtTestDataSheet.getRow(i);
                	
                	try{
                    	if (Cell.CELL_TYPE_STRING == row.getCell(0).getCellType()){
                    		//String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
                    		String strCellValueVFPage = row.getCell(0).getStringCellValue().trim();
                    			                    	
                    		if (!strCellValueVFPage.equals("")){
                    			// if First Col is VFPage-End then the rest 2 col will be empty (exception for blank cell)
                        		if (!strCellValueVFPage.equalsIgnoreCase("ALL Profiles-End")){                    			

                    				for (String strProfile : alALLProfilesinSFDC){
                    					ArrayList <String> alfinalArray = new ArrayList <String> ();
                    					ArrayList <String> alAppendArrayToExistingEntries = new ArrayList <String> ();
                    					alAppendArrayToExistingEntries = hmapProfileVFPageAccess.get(strProfile);
                    					
                    					if (alAppendArrayToExistingEntries.size() > 0){
                    						for (String str : alAppendArrayToExistingEntries){
                    							//str = str + "#" + strCellValueAccess;
                    							alfinalArray.add(str);
                    						}
                    					}
                    					alfinalArray.add(strCellValueVFPage);                						
                    					hmapProfileVFPageAccess.put(strProfile, alfinalArray); 
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
                
        		hmapProfileVFPageAccess.put( "Standard", hmapProfileVFPageAccess.remove( "Standard User" ) );
        		hmapProfileVFPageAccess.put( "Admin", hmapProfileVFPageAccess.remove( "System Administrator" ) );
        		hmapProfileVFPageAccess.put( "ContractManager", hmapProfileVFPageAccess.remove( "Contract Manager" ) );
        		hmapProfileVFPageAccess.put( "MarketingProfile", hmapProfileVFPageAccess.remove( "Marketing User" ) );
        		hmapProfileVFPageAccess.put( "ReadOnly", hmapProfileVFPageAccess.remove( "Read Only" ) );
        		hmapProfileVFPageAccess.put( "SolutionManager", hmapProfileVFPageAccess.remove( "Solution Manager" ) );
        		hmapProfileVFPageAccess.put( "StandardAul", hmapProfileVFPageAccess.remove( "Standard Platform User" ) );
        		
                Set set = hmapProfileVFPageAccess.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			System.out.print("PROFILE: "+ mentry.getKey() + ".... COMBINATIONS: ");
        			System.out.println(mentry.getValue());
        			FileOperations.writeToLog("-------------------------------------------------");
        			FileOperations.writeToLog("Profile: "+ mentry.getKey());
        			FileOperations.writeToLog("VFPage: "+ mentry.getValue().toString());
        		}
            }
            else{
            	//System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","Profile/VF Page Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
                e.printStackTrace();
                FileOperations.writeToLog("Inside Catch 2: " +e.getMessage());
        }
	}

}
