package Test;

import static GUI.PostDepTool.hmapProfileFieldObjAccess;
import static GUI.PostDepTool.hmapProfileObjectAccessRecTypAssDefault;
import static Utilities.ExcelPOI.strTestDataFilePath;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sforce.soap.enterprise.EnterpriseConnection;
import com.sforce.soap.enterprise.LoginResult;
import com.sforce.soap.metadata.AsyncResult;
import com.sforce.soap.metadata.CustomObject;
import com.sforce.soap.metadata.DescribeMetadataObject;
import com.sforce.soap.metadata.DescribeMetadataResult;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.PermissionSet;
import com.sforce.soap.metadata.PermissionSetFieldPermissions;
import com.sforce.soap.metadata.Profile;
import com.sforce.soap.metadata.ProfileApexClassAccess;
import com.sforce.soap.metadata.ProfileApexPageAccess;
import com.sforce.soap.metadata.ProfileFieldLevelSecurity;
import com.sforce.soap.metadata.ProfileLayoutAssignment;
import com.sforce.soap.metadata.ProfileObjectPermissions;
import com.sforce.soap.metadata.ProfileRecordTypeVisibility;
import com.sforce.soap.metadata.ReadResult;
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.metadata.UpdateMetadata_element;
import com.sforce.soap.partner.Connector;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.partner.QueryResult;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;
import com.sun.xml.internal.ws.util.MetadataUtil;

import Utilities.ExcelPOI;
import Utilities.FileOperations;
public class MetadataAPITest {
	public static MetadataConnection con = null;
	//public static MetadataConnection con = null;
	public static PartnerConnection con2 = null;
	public static File file;
	public static Map<String, ArrayList<String>> hmapProfileFieldObjAccess;
	
	public MetadataAPITest() throws IOException{
	}
		
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		/**
		 * Login utility.
		 */
		//login();
		//browseForExcelFile();		
		/*createFile();
		
		ExcelPOI.strTestDataFilePath = "C:\\Users\\267567\\Desktop\\Salesforce Release 39.xlsx";
		
		// Get starting row to read frm RTP Excel (After All Profiles)
		int iRowAllProfileEnds = -1, iStartRow = -1;
				
		iRowAllProfileEnds = ExcelPOI.GetRowIndexofValueinCol("FLS_PostDepTool", "ALL Profiles End", 0, 0);
		if (iRowAllProfileEnds != -1)
			iStartRow = ExcelPOI.GetRowIndexofValueinCol("FLS_PostDepTool", "Field", 0, iRowAllProfileEnds);
		
		System.out.println(iRowAllProfileEnds);
		System.out.println(iStartRow);
		
		// get the list of Unique profiles from the Profile column in RTP
		ArrayList <String>alProfilesfrmRTP = new ArrayList();
		if (iStartRow != -1){
			alProfilesfrmRTP = ExcelPOI.GetUniqueRowsfromColumn("FLS_PostDepTool", "Profile", iStartRow);
			if (alProfilesfrmRTP.contains("Profile"))
				alProfilesfrmRTP.remove("Profile");
			
			for (String str : alProfilesfrmRTP){
				//System.out.println(str);
			}
		}
		
		// Retrieve Field Object AccessLevel for the Profiles in RTP
		getFieldAPIObjectAPIAccessLevelforProfile (alProfilesfrmRTP, iStartRow);
		
		//describeMetadata();
		//setFLSforProfile("Design Support Engineer", "Customer_Survey_Result__c", "Additional_Feedback__c");
		
		Set set = hmapProfileFieldObjAccess.entrySet();
		Iterator iterator = set.iterator();
		while(iterator.hasNext()) {
			Map.Entry mentry = (Map.Entry)iterator.next();
			//System.out.print("key is: "+ mentry.getKey() + " & Value is: ");
			//System.out.println(mentry.getValue());
			writeToLog("key is: "+ mentry.getKey() + " & Value is: ");
			writeToLog(mentry.getValue().toString());
			
			String strProfile = mentry.getKey().toString();
			ArrayList <String> alFieldObjectAccess = new ArrayList();
			alFieldObjectAccess = (ArrayList<String>) mentry.getValue();
			createProfileFieldLvlSecurityArray (strProfile, alFieldObjectAccess);
			
			
			
		}*/
		
		login();
		//readAccess();
		setVFPageAccessforProfile();
		//setFLSforProfile();
		
	}
	
	public static boolean connectToSFDCDB (String strUsername, String strPasswd, String strSecurityToken){
        boolean iConnStatus = true;
        try{
            ConnectorConfig config = new ConnectorConfig();
            
            System.out.println("Username: "+strUsername);
            System.out.println("PWD: "+strPasswd);
            config.setUsername(strUsername);
            config.setPassword(strPasswd + strSecurityToken);
            
            String URL = "Sandbox";
            String authEndPoint = "";
            if (URL.equalsIgnoreCase("Sandbox"))
            	authEndPoint =  "https://test.salesforce.com/services/Soap/u/29.0";
			else if (URL.equalsIgnoreCase("PROD"))
				authEndPoint =  "https://login.salesforce.com/services/Soap/u/29.0";

            config.setAuthEndpoint(authEndPoint);
            
            con2 = Connector.newConnection(config);
            // display some current settings
            System.out.println("Auth EndPoint: "+config.getAuthEndpoint());
            System.out.println("Service EndPoint: "+config.getServiceEndpoint());
            System.out.println("Username: "+config.getUsername());
            System.out.println("SessionId: "+config.getSessionId());
                      
            return iConnStatus;
        } 
        catch(Exception e){
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,"Error in connecting to Salesforce","SFDC Connection Error",JOptionPane.ERROR_MESSAGE);
            iConnStatus = false;
            return iConnStatus;
        }
    }
	
	public static MetadataConnection login() throws ConnectionException {
        final String USERNAME = "sysuser@sunpower.com.uat"; //testdeploy";
        final String PASSWORD = "Solar123"; //925i5XnTrtXmrVW2seIoeIke";
        final String URL = "https://test.salesforce.com/services/Soap/c/29.0"; // "https://login.salesforce.com/services/Soap/c/29.0"; 
        		//"https://test.salesforce.com/services/Soap/c/37.0";
        System.out.println("Logging in as: " +USERNAME + " in Env: " +URL);
        final LoginResult loginResult = loginToSalesforce(USERNAME, PASSWORD, URL);
        System.out.println("Logged in..");
        return createMetadataConnection(loginResult);
    }
		 
    private static MetadataConnection createMetadataConnection(final LoginResult loginResult) throws ConnectionException {
        final ConnectorConfig config = new ConnectorConfig();
        config.setServiceEndpoint(loginResult.getMetadataServerUrl());
        config.setSessionId(loginResult.getSessionId());
        con = new MetadataConnection(config);
        return con;
    }
		    
    private static LoginResult loginToSalesforce(final String username, final String password, final String loginUrl) throws ConnectionException {
        final ConnectorConfig config = new ConnectorConfig();
        config.setAuthEndpoint(loginUrl);
        config.setServiceEndpoint(loginUrl);
        config.setManualLogin(true);
        return (new EnterpriseConnection(config)).login(username, password); //+"pTQo7s8YWGi2CPjqLTkRqmGB"
    }
    
    public static void describeMetadata() {
	  try {
	    double apiVersion = 37.0;
	    // Assuming that the SOAP binding has already been established.
	    DescribeMetadataResult res = con.describeMetadata(apiVersion);
	    //StringBuffer sb = new StringBuffer();
	    if (res != null && res.getMetadataObjects().length > 0) {
	      for (DescribeMetadataObject obj : res.getMetadataObjects()) {
	    	  System.out.println("***************************************************");
	    	  System.out.println("XMLName: " + obj.getXmlName());
	    	  System.out.println("DirName: " + obj.getDirectoryName());
	    	  System.out.println("Suffix: " + obj.getSuffix());
	    	  System.out.println("***************************************************");
	      }
	    } else {
	    	System.out.println("Failed to obtain metadata types.");
	    }
	    //System.out.println(sb.toString());
	  } catch (ConnectionException ce) {
	    ce.printStackTrace();
	  }
	}
    
    public static void browseForExcelFile(){
    	// Browse for the Excel file
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx");
        fileChooser.setFileFilter(filter);

        fileChooser.setDialogTitle("Select the RTP sheet");
        int userSelection = fileChooser.showDialog(null, "Select Excel");

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fExcelFile = fileChooser.getSelectedFile();
            ExcelPOI.strTestDataFilePath = fExcelFile.getAbsolutePath();
            System.out.println("Excel file: " + fExcelFile.getAbsolutePath());
        }
    }
    
    public static void getFieldAPIObjectAPIAccessLevelforProfile (ArrayList<String> alProfilesfrmRTP, int iStartRow){
        String strSheetName = "FLS_PostDepTool";
        String strColNameToRead_1 = "Profile";
        String strColNameToRead_2 = "API Name";
        String strColNameToRead_3 = "Object";
        String strColNameToRead_4 = "Access Level";
        int iRowToStartReadingFrom = iStartRow;
                      
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
            	for (String strProfile : alProfilesfrmRTP){
            		hmapProfileFieldObjAccess.put(strProfile, new ArrayList<String>());
            	}
            	
            	for (i = iRowToStartReadingFrom+1; i <= shtTestDataSheet.getLastRowNum(); i++) {
                    row = shtTestDataSheet.getRow(i);
                    try{
	                    if (Cell.CELL_TYPE_STRING == row.getCell(j).getCellType()){
	                    	String strCellValueProfile = row.getCell(j).getStringCellValue().trim();
	                    	String strCellValueFieldAPI = row.getCell(k).getStringCellValue().trim();
	                    	String strCellValueObject = row.getCell(l).getStringCellValue().trim();
	                    	String strCellValueAccess = row.getCell(m).getStringCellValue().trim();
	                    	
	                    	if (!strCellValueProfile.equals("")){
	                    		if (!strCellValueProfile.equals("Profile")){
	                    			ArrayList<String> alFieldObjectAccess = hmapProfileFieldObjAccess.get(strCellValueProfile);
	                    			String strFieldObjAccess = strCellValueObject + "." + strCellValueFieldAPI + "|" + strCellValueAccess;
	                    			//System.out.println(strFieldObjAccess); 
	                    			alFieldObjectAccess.add(strFieldObjAccess);
	                    			hmapProfileFieldObjAccess.put(strCellValueProfile, alFieldObjectAccess);                    			
	                    		}
	                    	}
	                    }
                    }catch(Exception e){
                    	//e.printStackTrace();
                    	//writeToLog ();
                    }
            	}
                //JOptionPane.showMessageDialog(null,"DONE!","DONE",JOptionPane.INFORMATION_MESSAGE);
            	
            	Set set = hmapProfileFieldObjAccess.entrySet();
        		Iterator iterator = set.iterator();
        		while(iterator.hasNext()) {
        			Map.Entry mentry = (Map.Entry)iterator.next();
        			//System.out.print("key is: "+ mentry.getKey() + " & Value is: ");
        			//System.out.println(mentry.getValue());
        			writeToLog("key is: "+ mentry.getKey() + " & Value is: ");
        			writeToLog(mentry.getValue().toString());
        		}
            }
            else{
                //System.out.println("Column Name not found");
                JOptionPane.showMessageDialog(null,"Column Missing!","Profile/API Name/Object/Access level Column not found in RTP Excel",JOptionPane.ERROR_MESSAGE);
            }
        }
        catch(Exception e){
            e.printStackTrace();
        }
    }
    
    public static void createFile() throws IOException{
    	DateFormat dateFormat = new SimpleDateFormat("MM/dd HH:mm:ss");
	    Date date = new Date();
    	String strFileName = "PostDepLog" + dateFormat.format(date) + ".txt";
    	strFileName = strFileName.replace(" ", "");
    	strFileName = strFileName.replace("/", "-");
    	strFileName = strFileName.replace(":", "-");
    	file = new File(strFileName);
        
        // creates the file
        file.createNewFile();
    }
    
    public static void writeToLog(String strContentToFile) throws IOException{
    	/*DateFormat dateFormat = new SimpleDateFormat("MM/dd HH:mm:ss");
	    Date date = new Date();
    	String strFileName = "PostDepLog" + dateFormat.format(date);
    	strFileName = strFileName.replace(" ", "");
    	strFileName = strFileName.replace("/", "-");
    	strFileName = strFileName.replace(":", "-");
    	File file = new File(strFileName);
        
        // creates the file
        file.createNewFile();*/
        
        // creates a FileWriter Object
        FileWriter writer = new FileWriter(file); 
        
        // Writes the content to the file
        //writer.write(strContentToFile); 
        writer.append(strContentToFile);
        writer.flush();
        writer.close();
    }
    
    public static void createProfileFieldLvlSecurityArray (String strProfile, ArrayList<String> alFieldObjectAccess) throws IOException, ConnectionException{
    	try{
	    	int iArraySize = alFieldObjectAccess.size();
	    	ProfileFieldLevelSecurity[] fieldPermissions = new ProfileFieldLevelSecurity[iArraySize];
	    	
	    	int iCounter = 0;
	    	for (String strFOA : alFieldObjectAccess){
				String strFieldObj = strFOA.split("|")[0];
				String strAccess = strFOA.split("|")[1];
				
				fieldPermissions[iCounter] = new ProfileFieldLevelSecurity();
		    	fieldPermissions[iCounter].setField(strFieldObj);
		    	
		    	if (strAccess.equalsIgnoreCase("Read ONLY"))
		    		fieldPermissions[iCounter].setEditable(false);
		    	else if (strAccess.equalsIgnoreCase("Read/Write"))
		    		fieldPermissions[iCounter].setEditable(true);
		    	else
		    		writeToLog("Error: Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
		    	
		    	fieldPermissions[iCounter].setReadable(true);
		    	iCounter ++;
			}
	    	
	    	//setFLSforProfile (strProfile, fieldPermissions);
    	}catch(Exception e){
    		writeToLog("Error in creating ProfileFieldLvlSecurityArray for Profile: " +strProfile);
    	}
    }
    
    public static void setFLSforProfile(/*String strProfile , ProfileFieldLevelSecurity[] fieldPermissions*/) throws ConnectionException{
    	try{
    		String strProfile = "Marketing Connect5";
	    	Profile prof = new Profile();
	    	
	    	ProfileFieldLevelSecurity[] fieldPermissions = new ProfileFieldLevelSecurity[1];
	    	fieldPermissions[0] = new ProfileFieldLevelSecurity();
	    	fieldPermissions[0].setField("Opportunity_Role__c.DirectMargin__c");
	    	//fieldPermissions[0].setEditable(false);
	    	fieldPermissions[0].setReadable(true);
	    	fieldPermissions[0].setEditable(true);
	    	
	    	/*ProfileFieldLevelSecurity pfs = new ProfileFieldLevelSecurity();
	    	pfs.setField(strObjectName + "." + strFieldName);
	    	//pfs.setReadable(true);
	    	pfs.setEditable(false);	    	
	    	fieldPermissions[0] = pfs;*/
	    	
	    	/*fieldPermissions[1] = new ProfileFieldLevelSecurity();
	    	fieldPermissions[1].setField("Opportunity.Acres_of_Trees__c");
	    	fieldPermissions[1].setEditable(false);
	    	fieldPermissions[0].setReadable(false);*/
	    	
	    	prof.setFullName(strProfile);
	    	prof.setFieldPermissions(fieldPermissions);
	    	
	    	SaveResult[] arsTab =  con.updateMetadata(new Metadata[] {prof});
	    	
	    	for (SaveResult r : arsTab) {
	            if (r.isSuccess()) {
	                System.out.println("Updated component: " + r.getFullName());
	            } else {
	                System.out
	                        .println("Errors were encountered while updating "
	                                + r.getFullName());
	                for (com.sforce.soap.metadata.Error e : r.getErrors()) {
	                    System.out.println("Error message: " + e.getMessage());
	                    System.out.println("Status code: " + e.getStatusCode());
	                }
	            }
	        }
    	}catch(ConnectionException ce){
    		ce.printStackTrace();
    	}
    }

    /*public static void createProfileFieldLvlSecurityArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileFieldObjAccess.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ProfileFieldLevelSecurity[] fieldPermissions = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alFieldObjectAccess = new ArrayList();
    			alFieldObjectAccess = (ArrayList<String>) mentry.getValue();
    			
    			int iArraySize = alFieldObjectAccess.size();
    			if (iArraySize > 0){
    				fieldPermissions = new ProfileFieldLevelSecurity[iArraySize];
    				
    				int iCounter = 0;
    		    	for (String strFOA : alFieldObjectAccess){
    					String strFieldObj = strFOA.split("#")[0];
    					String strAccess = strFOA.split("#")[1];
    					
    					fieldPermissions[iCounter] = new ProfileFieldLevelSecurity();
    			    	fieldPermissions[iCounter].setField(strFieldObj);
    			    	
    			    	//Read ONLY: setEditable(false).. Read/Write: setEditable(true)
    			    	if (strAccess.equalsIgnoreCase("Read ONLY"))
    			    		fieldPermissions[iCounter].setEditable(false); 
    			    	else if (strAccess.equalsIgnoreCase("Read/Write"))
    			    		fieldPermissions[iCounter].setEditable(true); 
    			    	else{
    			    		System.out.println("Error: Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
    			    		FileOperations.writeToLog("Error: Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
    			    		JOptionPane.showMessageDialog(null,"Error: Access Level provided in excel - " +strAccess + "for Profile: " +strProfile,"Wrong FLS Access",JOptionPane.ERROR_MESSAGE);
    			    	}
    			    	
    			    	fieldPermissions[iCounter].setReadable(true);
    			    	iCounter ++;
    				}
    		    	
    		    	if (iCounter > 0){
	    		    	strListOfProf = strListOfProf + strProfile + "|";
	    		    	Profile prof = new Profile();
	    		    	prof.setFullName(strProfile);
	    		    	prof.setFieldPermissions(fieldPermissions);
	    		    	
	    		    	arrayOf10Profiles.add(prof);
	    		    	iCountOfProfilesNotMoreThan10 ++;
	    		    	//System.out.println(strProfile);
    		    	}
    		    	
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		FileOperations.writeToLog("-------------------------------------------------");
    			    	FileOperations.writeToLog("Setting FLS for Profiles: " +strListOfProf);
    		    		System.out.println("-------------------------------------------------");
    			    	System.out.println("Setting FLS for Profiles: " +strListOfProf);	 
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    			    	
    			    	updateMetadataFor10Profiles (arrayOf10Profiles);
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			FileOperations.writeToLog("-------------------------------------------------");
		    	FileOperations.writeToLog("Setting FLS for Profiles: " +strListOfProf);
	    		System.out.println("-------------------------------------------------");
		    	System.out.println("Setting FLS for Profiles: " +strListOfProf);
		    	updateMetadataFor10Profiles (arrayOf10Profiles);
    		}
	    	
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileFieldLvlSecurityArray for Profile: " +strProfile);
    	}
    }*/
    
    public static void readAccess () throws ConnectionException{
    	
    	try{
    		connectToSFDCDB("sysuser@sunpower.com.testdeploy", "Solar123", "");
    		//String strSOQLQuery = "SELECT PermissionsEdit FROM ObjectPermissions WHERE SobjectType = 'pca__Action__c' and parentid in (select id from permissionset where PermissionSet.Profile.Name = 'SunPower Administrator')";
    		//String strSOQLQuery = "SELECT PermissionsRead FROM ObjectPermissions WHERE SobjectType = 'Alliance_Partner__c' and parentid in (select id from permissionset where PermissionSet.Profile.Name = 'Partner TPS')";
    		//String strSOQLQuery = "SELECT PermissionsEdit FROM FieldPermissions WHERE Field = 'Opportunity.Final_Amount__c' AND SobjectType = 'Opportunity' AND parentid in (select id from permissionset where PermissionSet.Profile.Name = 'Admin')";
    		String strSOQLQuery = "SELECT PermissionsDelete FROM ObjectPermissions WHERE SobjectType = 'Plan_Type__c' and parentid in (select id from permissionset where PermissionSet.Profile.Name = 'Admin')";
    		QueryResult qr = con2.query(strSOQLQuery);
    		System.out.println(qr.getSize());
    		System.out.println(qr.hashCode());
    		//System.out.println(qr.getRecords()[0].getField("PermissionsEdit"));
    		
    		/*if( qr.equals(true))
    			System.out.println("true");
    		else
    			System.out.println("false");*/
    		System.out.println(qr.toString());
    		/*String strProfile = "SunPower Administrator";
	    	Profile prof = new Profile();
	    	
	    	ProfileObjectPermissions[] profObjPer = new ProfileObjectPermissions[1];
	    	profObjPer[0] = new ProfileObjectPermissions();
	    	profObjPer[0].setObject("pca__Action__c");
	    	
	    	prof.setFullName(strProfile);
	    	
	    	ReadResult readResult = con.readMetadata("CustomObject", new String[] {"pca__Action__c"});
	    	Metadata[] mdInfo = readResult.getRecords();
   	        System.out.println("Number of component info returned: " + mdInfo.length);
	    	
   	        for (Metadata md : mdInfo) {
   	        	if (md != null) {
   	        		CustomObject obj = (CustomObject) md;
   	        		//obj.
   	        	}
   	        }
	    	
	    	System.out.println(profObjPer[0].getAllowCreate());
	    	System.out.println(profObjPer[0].getAllowDelete());
	    	System.out.println(profObjPer[0].getAllowEdit());
	    	System.out.println(profObjPer[0].getAllowRead());
	    	System.out.println(profObjPer[0].getViewAllRecords());
	    	System.out.println(profObjPer[0].getModifyAllRecords());*/
    	}catch(ConnectionException ce){
			ce.printStackTrace();
		}
    	
    }
    	
    public static void setVFPageAccessforProfile () throws ConnectionException{
    	
    	try{
    		//readAccess();
    		
	    	String strProfile = "ReadOnly";
	    	Profile prof = new Profile();
	    	
	    	/*ProfileApexPageAccess[] profVF = new ProfileApexPageAccess[1];
	    	profVF[0] = new ProfileApexPageAccess();
	    	profVF[0].setApexPage("ACBParanetLink");
	    	profVF[0].setEnabled(true);
	    	
	    	ProfileApexClassAccess[] profApClass = new ProfileApexClassAccess[1];
	    	profApClass[0] = new ProfileApexClassAccess();
	    	profApClass[0].setApexClass("ACBParanetLinkController");
	    	profApClass[0].setEnabled(true);*/
	    	
	    	ProfileObjectPermissions[] profObjPer = new ProfileObjectPermissions[1];
	    	profObjPer[0] = new ProfileObjectPermissions();
	    	profObjPer[0].setObject("StandardLineItem__c");
	    	
	    	profObjPer[0].setAllowRead(true);
	    	profObjPer[0].setAllowCreate(false);
	    	profObjPer[0].setAllowDelete(false);
	    	profObjPer[0].setAllowEdit(false);
	    	profObjPer[0].setViewAllRecords(false);
	    	profObjPer[0].setModifyAllRecords(false);
	    	
	    	/*ProfileRecordTypeVisibility[] profReTypVis = new ProfileRecordTypeVisibility[1];
	    	profReTypVis[0] = new ProfileRecordTypeVisibility();
	    	profReTypVis[0].setRecordType("Design__c.Proposal_Design");
	    	profReTypVis[0].setDefault(true);
	    	profReTypVis[0].setVisible(true);*/
	    	   	
	    	/*ProfileLayoutAssignment[] profReTypVis = new ProfileLayoutAssignment[1];
	    	profReTypVis[0] = new ProfileLayoutAssignment();
	    	//profReTypVis[0].setRecordType("Document__c.TPS");
	    	profReTypVis[0].setLayout("Document__c-Partner TPS Document Layout");*/
	    	
	    	/*PermissionSetFieldPermissions[] fieldPermissions = new PermissionSetFieldPermissions[1];
	    	fieldPermissions[0] = new PermissionSetFieldPermissions();
	    	fieldPermissions[0].setField("Residential_Project__c.Project_Task_Max_End_Date__c");
	    	fieldPermissions[0].setEditable(true);
	    	fieldPermissions[0].setReadable(true);
	    	
	    	PermissionSet permSet = new PermissionSet();		    	
    		permSet.setFullName("Access_to_Account_Fields");	
    		permSet.setFieldPermissions(fieldPermissions);*/
    		
	    	prof.setFullName(strProfile);
	    	prof.setObjectPermissions(profObjPer);
	    	//prof.setRecordTypeVisibilities(profReTypVis);
	    	//prof.setPageAccesses(profVF);
	    	//prof.setClassAccesses(profApClass);
	    	//prof.setLayoutAssignments(profReTypVis);
	    	
	    	SaveResult[] arsTab =  con.updateMetadata(new Metadata[] {prof});
	    	
	    	for (SaveResult r : arsTab) {
	            if (r.isSuccess()) {
	                System.out.println("Updated component: " + r.getFullName());
	            } else {
	                System.out
	                        .println("Errors were encountered while updating "
	                                + r.getFullName());
	                for (com.sforce.soap.metadata.Error e : r.getErrors()) {
	                    System.out.println("Error message: " + e.getMessage());
	                    System.out.println("Status code: " + e.getStatusCode());
	                }
	            }
	        }
	    	//System.out.println(prof.getObjectPermissions());
	    	
	    	/*System.out.println(profObjPer[0].getAllowCreate());
	    	System.out.println(profObjPer[0].getAllowDelete());
	    	System.out.println(profObjPer[0].getAllowEdit());
	    	System.out.println(profObjPer[0].getAllowRead());
	    	System.out.println(profObjPer[0].getViewAllRecords());
	    	System.out.println(profObjPer[0].getModifyAllRecords());*/
		}catch(ConnectionException ce){
			ce.printStackTrace();
		}
    	
    }

}
