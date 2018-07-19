package Components;

import com.sforce.soap.enterprise.EnterpriseConnection;
import com.sforce.soap.enterprise.LoginResult;
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
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.partner.Connector;
import com.sforce.soap.partner.PartnerConnection;
import com.sforce.soap.partner.QueryResult;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;

import static GUI.PostDepTool.alALLProfilesinSFDC;
import static GUI.PostDepTool.hmapProfileFieldObjAccess;
import static GUI.PostDepTool.hmapProfileVFPageAccess;
import static GUI.PostDepTool.hmapProfileApexClassAccess;
import static GUI.PostDepTool.hmapProfileObjectAccessRecTypAssDefault;
import static GUI.PostDepTool.hmapProfileObjectRecTypePageLayout;
import static GUI.PostDepTool.hmapPermissionSetFieldObjAccess;

import Utilities.ExcelPOI;
import Utilities.FileOperations;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JOptionPane;

public class MetaDataAPI_FLScheck { //change class name to MetaDataAPI later as reqd
	public static MetadataConnection con = null;
	public static PartnerConnection con2 = null;
	
	public static boolean loginUsingPartnerConnection (){
        boolean iConnStatus = true;
        try{
        	final String USERNAME = ExcelPOI.ReadDataFromExcel("Types", "LoginID", 1);
			final String PASSWORD = ExcelPOI.ReadDataFromExcel("Types", "Pass", 1);
			String URL = ExcelPOI.ReadDataFromExcel("Types", "Env", 1);
			
            ConnectorConfig config = new ConnectorConfig();
            
            System.out.println("Username: "+USERNAME);
            System.out.println("PWD: "+PASSWORD);
            config.setUsername(USERNAME);
            config.setPassword(PASSWORD); // + strSecurityToken);
            
            String authEndPoint = "";
            if (URL.equalsIgnoreCase("Sandbox"))
            	authEndPoint =  "https://test.salesforce.com/services/Soap/u/29.0";
			else if (URL.equalsIgnoreCase("PROD"))
				authEndPoint =  "https://login.salesforce.com/services/Soap/u/29.0";

            config.setAuthEndpoint(authEndPoint);
            
            System.out.println("SOAP PARTNER Connection: Logging in as: " +USERNAME + " in Env: " +URL);
            con2 = Connector.newConnection(config);
            System.out.println("Logged in...");
            // display some current settings
            /*System.out.println("Auth EndPoint: "+config.getAuthEndpoint());
            System.out.println("Service EndPoint: "+config.getServiceEndpoint());
            System.out.println("Username: "+config.getUsername());
            System.out.println("SessionId: "+config.getSessionId());*/
                      
            return iConnStatus;
        } 
        catch(Exception e){
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,"Error in connecting to Salesforce","SFDC Connection Error",JOptionPane.ERROR_MESSAGE);
            iConnStatus = false;
            return iConnStatus;
        }
    }
	
	public static MetadataConnection login() throws ConnectionException, IOException {
		try{
			final String USERNAME = ExcelPOI.ReadDataFromExcel("Types", "LoginID", 1);
			final String PASSWORD = ExcelPOI.ReadDataFromExcel("Types", "Pass", 1);
			String URL = ExcelPOI.ReadDataFromExcel("Types", "Env", 1);
			
			System.out.println("METADATA Connection: Logging in as: " +USERNAME + " in Env: " +URL);
			
			if (URL.equalsIgnoreCase("Sandbox"))
				URL =  "https://test.salesforce.com/services/Soap/c/29.0";
			else if (URL.equalsIgnoreCase("PROD")){
				URL =  "https://login.salesforce.com/services/Soap/c/29.0";
				JOptionPane.showMessageDialog(null,"Env: PROD, RTP: "+ ExcelPOI.strTestDataFilePath, "RUNNING IN PRODUCTION !!",JOptionPane.ERROR_MESSAGE);
			}
			
	        //final String USERNAME = "subhra.bikashdas@sunpowercorp.com.stg";
	        //final String PASSWORD = "aaaa1111"; //l6NKlfiLcBjpNppeOt1sCDdq"; //zIicVrPIZqc827XpJVak7TCN";
	        //final String URL =  "https://test.salesforce.com/services/Soap/c/29.0"; //"https://test.salesforce.com/services/Soap/c/38.0/0DF220000004CHW";
	        		//"https://test.salesforce.com/services/Soap/c/37.0";
			
	        final LoginResult loginResult = loginToSalesforce(USERNAME, PASSWORD, URL);
	        System.out.println("Logged in..");
	        
	        return createMetadataConnection(loginResult);
		}catch (Exception e){
			e.printStackTrace();
			return null;
		}
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
    
    public static void getListofAllProfilesinSFDC (){
    	//alALLProfilesinSFDC = new ArrayList <String> ();
    	/*alALLProfilesinSFDC.add("SunPower User");
    	alALLProfilesinSFDC.add("Prof2");
    	alALLProfilesinSFDC.add("Prof3");*/
    	
    	alALLProfilesinSFDC = ExcelPOI.GetAllRowsfromColumn("Types", "All Profiles");
    }
    
    public static void getListofAllPermissionSetsinSFDC (){
    	
    }
   
    public static void createProfileFieldLvlSecurityArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileFieldObjAccess.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ArrayList <ProfileFieldLevelSecurity> alFieldPermissions = null;
    		ProfileFieldLevelSecurity[] fieldPermissions = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		
    		System.out.println("****************************************************************");
    		System.out.println("        Checking FLS settings in salesforce for Profiles:       ");
    		System.out.println("****************************************************************");
    		FileOperations.writeToLog("****************************************************************");
    		FileOperations.writeToLog("        Checking FLS settings in salesforce for Profiles:       ");
    		FileOperations.writeToLog("****************************************************************");
    		
    		int iProfileCount = 1; //just for the log
    		while(iterator.hasNext()) { //parse through all the profiles one by one
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alFieldObjectAccess = new ArrayList();
    			alFieldObjectAccess = (ArrayList<String>) mentry.getValue(); // get all the "Obj # AccessLvl # RecType # Default" combinations for the profile
    			
    			System.out.println(iProfileCount + ".. " + strProfile);
    			FileOperations.writeToLog(iProfileCount + ".. " + strProfile);
    			iProfileCount++;
    			
    			int iArraySize = alFieldObjectAccess.size();
    			if (iArraySize > 0){ // if there are combinations of "Obj # Access Lvl # RecType # Default" for the profile
    				alFieldPermissions = new ArrayList <ProfileFieldLevelSecurity> ();
    				fieldPermissions = new ProfileFieldLevelSecurity[iArraySize];
    				
    				for (String strFOA : alFieldObjectAccess){ // parse through all the "Obj # Access Lvl # RecType # Default" combinations one by one
    					String strFieldObj = strFOA.split("#")[0];
    					
    					// String strObj = strFieldObj.split(".")[0]; split func not working for dot, hence replacing dot by #
    					strFieldObj = strFieldObj.replace(".", "#");
    					String strObj = strFieldObj.split("#")[0];
    					strFieldObj = strFieldObj.replace("#", "."); // replacing # by dot again since obj.field format is used later in fieldPermission.setField(strFieldObj);
    					
    					String strAccess = strFOA.split("#")[1];
    					    					
    					String strSOQLQuery = "";
    					String strPermission = "PermissionsEdit";
    					String strPermissionVaue = "";
    					
    					if (!strAccess.equals("")){	
    						ProfileFieldLevelSecurity fieldPermission = new ProfileFieldLevelSecurity();
    						fieldPermission.setField(strFieldObj);
    						
    						//System.out.print(strFieldObj + ".. "); //Checking FLS
    						
    						// Check if Query return any value. If size = 0, then thr is no Read Access for the obj in SFDC.
    						strSOQLQuery = "SELECT " +strPermission+ " FROM FieldPermissions WHERE Field = '" +strFieldObj +"' and SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    						QueryResult qr = con2.query(strSOQLQuery);
    						int iRecSize = qr.getSize();
    						
    						// iRecSize = 0 if Field is NOT VISIBLE. And if Field is NOT VISIBLE, Read/Write or ReadOnly Access is N/A
    						Boolean bAccessLvinRTPGr8rThnAccessLvlinSFDC = true;
    						String strErrMsg = "";
    						if (iRecSize > 0){
    					 		//SFDC has Read/Write Access for Obj. RTP has Read ONLY for Obj
    					 		if (strAccess.equalsIgnoreCase("Read ONLY")){
    					 			strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    					 			if (strPermissionVaue.equals("true")){
    					 				bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    					 				strErrMsg = "Error: FLS given in salesforce: Read/Write BUT in RTP: Read ONLY for Obj.Field: " +strFieldObj+ " Prof: " +strProfile;
    					 				System.out.println(strErrMsg);
    					 				FileOperations.writeToLog(strErrMsg);
    					 			}
    					 			else{
    									fieldPermission.setReadable(true);
    									fieldPermission.setEditable(false); // redundant since it is already Read ONLY in the salesforce
    								}
    					 		}
    					 		else if (strAccess.equalsIgnoreCase("Read/Write")){
    					 			fieldPermission.setReadable(true);
    					 			fieldPermission.setEditable(true);
    					 		}
    					 		else if (strAccess.equalsIgnoreCase("No Access")){
    					 			fieldPermission.setReadable(false);
    					 		}
    	    			    	else{
    	    			    		System.out.println("Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
    	    			    		FileOperations.writeToLog("Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
    	    			    		JOptionPane.showMessageDialog(null,"Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile,"Wrong FLS Access",JOptionPane.ERROR_MESSAGE);
    	    			    		bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    	    			    	}
    						
    					 	}
    						else{ // This block would be encountered if the field has No Access (Not Readable) in Salesforce. In that case we need to give the necessary access that is given in RTP.
    							if (!strAccess.equalsIgnoreCase("No Access")){
    								//System.out.println("Alert: Field was in No Visible Status in salesforce. Now it is being set to visible status");
	    							fieldPermission.setReadable(true);
	    							if (strAccess.equalsIgnoreCase("Read/Write"))
	    								fieldPermission.setEditable(true);
	    							else if (strAccess.equalsIgnoreCase("Read ONLY"))
	    								fieldPermission.setEditable(false);
	    							else{
	    	    			    		System.out.println("Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
	    	    			    		FileOperations.writeToLog("Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile);
	    	    			    		JOptionPane.showMessageDialog(null,"Error: Wrong Access Level provided in excel - " +strAccess + "for Profile: " +strProfile,"Wrong FLS Access",JOptionPane.ERROR_MESSAGE);
	    	    			    		bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
	    	    			    	}
    							}
    							else{
    								fieldPermission.setReadable(false); // redundant, This -> bAccessLvinRTPGr8rThnAccessLvlinSFDC = false; would hav worked as well, which means it wont be conisdered for update, which is fine since it is already in No Access in salesforce
    							}
    						}
    						
    						if (bAccessLvinRTPGr8rThnAccessLvlinSFDC){
    							alFieldPermissions.add(fieldPermission);
    						}
    						
    					} //if (!strAccessLvl.equals("")){		    	
    					else{
    						//No access level set in RTP.
							System.out.println("Error: No Access Level set in RTP");
							FileOperations.writeToLog("Error: No Access Level set in RTP");
							JOptionPane.showMessageDialog(null,"Error: No Access Level set in RTP","FLS Error",JOptionPane.ERROR_MESSAGE);
    					}
    				}
    		    	
    			    if (alFieldPermissions.size() > 0){
    			    	strListOfProf = strListOfProf + strProfile + "|";
        		    	Profile prof = new Profile();
        		    	prof.setFullName(strProfile);
        		    	
        		    	fieldPermissions = new ProfileFieldLevelSecurity [alFieldPermissions.size()];
        		    	alFieldPermissions.toArray(fieldPermissions);
        		    	prof.setFieldPermissions(fieldPermissions);
        		    	
    			    	arrayOf10Profiles.add(prof);
    			    	iCountOfProfilesNotMoreThan10 ++;
    			    	//System.out.println(strProfile);
    			    }
    		    	
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		//FileOperations.writeToLog("-------------------------------------------------");
    			    	//FileOperations.writeToLog("Setting FLS for Profiles: " +strListOfProf);
    		    		//System.out.println("-------------------------------------------------");
    			    	//System.out.println("Setting FLS for Profiles: " +strListOfProf);	 
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    			    	
    			    	//updateMetadataFor10Profiles (arrayOf10Profiles, "FLS ACCESS");
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			else
    				FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			//FileOperations.writeToLog("-------------------------------------------------");
		    	//FileOperations.writeToLog("Setting FLS for last few Profiles: " +strListOfProf);
	    		//System.out.println("-------------------------------------------------");
		    	//System.out.println("Setting FLS for last few Profiles: " +strListOfProf);
		    	//updateMetadataFor10Profiles (arrayOf10Profiles, "FLS ACCESS");
    		}
	    		
    		System.out.println("************************ DONE **********************************");
    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileFieldLvlSecurityArray for Profile: " +strProfile);
    		System.out.println("Error in creating ProfileFieldLvlSecurityArray for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating ProfileFieldLvlSecurityArray for Profile: " +strProfile,"FLS Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void createProfileApexPageAccessArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileVFPageAccess.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ProfileApexPageAccess[] vfPageAccess = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alVFPage = new ArrayList();
    			alVFPage = (ArrayList<String>) mentry.getValue();
    			
    			int iArraySize = alVFPage.size();
    			if (iArraySize > 0){
    				vfPageAccess = new ProfileApexPageAccess[iArraySize];
    				
    				int iCounter = 0;
    				for (String strFOA : alVFPage){
    		    		vfPageAccess[iCounter] = new ProfileApexPageAccess();
    		    		vfPageAccess[iCounter].setApexPage(strFOA);
    		    		vfPageAccess[iCounter].setEnabled(true);
    			    	
    			    	iCounter ++;
    				}
    			    		
    				if (iCounter > 0){
	    				strListOfProf = strListOfProf + strProfile + "|";
	    		    	Profile prof = new Profile();
	    		    	prof.setFullName(strProfile);
	    		    	prof.setPageAccesses(vfPageAccess);
	    		    	
	    		    	arrayOf10Profiles.add(prof);
	    		    	iCountOfProfilesNotMoreThan10 ++;
	    		    	//System.out.println(strProfile);
    				}
    				
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		FileOperations.writeToLog("-------------------------------------------------");
    			    	FileOperations.writeToLog("Setting VF Page Access for Profiles: " +strListOfProf);
    		    		System.out.println("-------------------------------------------------");
    			    	System.out.println("Setting VF Page Access for Profiles: " +strListOfProf);
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    			    	
    			    	updateMetadataFor10Profiles (arrayOf10Profiles, "VF PAGE ACCESS");
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			else
    				FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			FileOperations.writeToLog("-------------------------------------------------");
		    	FileOperations.writeToLog("Setting VF Page Access for last few Profiles: " +strListOfProf);
	    		System.out.println("-------------------------------------------------");
		    	System.out.println("Setting VF Page Access for last few Profiles: " +strListOfProf);
		    	updateMetadataFor10Profiles (arrayOf10Profiles, "VF PAGE ACCESS");
    		}
	    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileApexPageAccessArray for Profile: " +strProfile);
    		System.out.println("Error in creating ProfileApexPageAccessArray for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating ProfileApexPageAccessArray for Profile: " +strProfile,"VF PAGE Access Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void createProfileApexClassAccessArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileApexClassAccess.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ProfileApexClassAccess[] apexClassAccess = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alApexClass = new ArrayList();
    			alApexClass = (ArrayList<String>) mentry.getValue();
    			
    			int iArraySize = alApexClass.size();
    			if (iArraySize > 0){
    				apexClassAccess = new ProfileApexClassAccess[iArraySize];
    				
    				int iCounter = 0;
    				for (String strFOA : alApexClass){
    					apexClassAccess[iCounter] = new ProfileApexClassAccess();
    					apexClassAccess[iCounter].setApexClass(strFOA);
    					apexClassAccess[iCounter].setEnabled(true);
    			    	
    			    	iCounter ++;
    				}
    		    	
    		    	if (iCounter > 0){
    		    		strListOfProf = strListOfProf + strProfile + "|";
        		    	Profile prof = new Profile();
        		    	prof.setFullName(strProfile);
        		    	prof.setClassAccesses(apexClassAccess);
        		    	
    		    		arrayOf10Profiles.add(prof);
    		    		iCountOfProfilesNotMoreThan10 ++;
    		    		//System.out.println(strProfile);
    		    	}
    		    	
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		FileOperations.writeToLog("-------------------------------------------------");
    			    	FileOperations.writeToLog("Setting Apex Class Access for Profiles: " +strListOfProf);
    			    	System.out.println("-------------------------------------------------");
    			    	System.out.println("Setting Apex Class Access for Profiles: " +strListOfProf);
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    			    	
    			    	updateMetadataFor10Profiles(arrayOf10Profiles, "APEX CLASS ACCESS");
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			else
    				FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			FileOperations.writeToLog("-------------------------------------------------");
		    	FileOperations.writeToLog("Setting Apex Class Access for last few Profiles: " +strListOfProf);
		    	System.out.println("-------------------------------------------------");
		    	System.out.println("Setting Apex Class Access for last few Profiles: " +strListOfProf);
		    	updateMetadataFor10Profiles(arrayOf10Profiles, "APEX CLASS ACCESS");
    		}
	    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileApexClassAccessArray for Profile: " +strProfile);
    		System.out.println("Error in creating ProfileApexClassAccessArray for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating ProfileApexClassAccessArray for Profile: " +strProfile,"APEX CLASS Access Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void createProfileObjPermAndProfileRecVisibilityArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileObjectAccessRecTypAssDefault.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ArrayList <ProfileObjectPermissions> alObjPermission = null;
    		ArrayList <ProfileRecordTypeVisibility> alProfRecTypVis = null;    	
    		ProfileObjectPermissions[] profObjPerms = null;
    		ProfileRecordTypeVisibility[] profRecTypVis = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		while(iterator.hasNext()) { //parse through all the profiles one by one
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alObjAccessLvlRecTypeDef = new ArrayList();
    			alObjAccessLvlRecTypeDef = (ArrayList<String>) mentry.getValue(); // get all the "Obj # AccessLvl # RecType # Default" combinations for the profile
    			
    			int iArraySize = alObjAccessLvlRecTypeDef.size();
    			if (iArraySize > 0){ // if there are combinations of "Obj # Access Lvl # RecType # Default" for the profile
    				alObjPermission = new ArrayList <ProfileObjectPermissions> ();
    				alProfRecTypVis = new ArrayList <ProfileRecordTypeVisibility> ();
    				
    				//int iCounter = 0;
    		    	//int iCounterObjRecTYpeAss = 0;
    				for (String strFOA : alObjAccessLvlRecTypeDef){ // parse through all the "Obj # Access Lvl # RecType # Default" combinations one by one
    					String strObj = strFOA.split("#")[0].trim();
    					String strAccessLvl = strFOA.split("#")[1].trim();
    					String strRecTypeAss = strFOA.split("#")[2].trim();
    					String strDefault = strFOA.split("#")[3].trim();
    					
    					String strSOQLQuery = "";
    					String strPermission = "PermissionsRead";
    					String strPermissionVaue = "";
    					
    					if (!strAccessLvl.equals("BLANK")){	
    						ProfileObjectPermissions objPermission = new ProfileObjectPermissions();
    						objPermission.setObject(strObj);
    						
    						//Trimming the spaces between the Read Create etc Access from RTP
    						List<String> arlistAccesslvl = Arrays.asList(strAccessLvl.split(","));
    						for(int i=0; i<arlistAccesslvl.size(); i++){
    							arlistAccesslvl.set(i, arlistAccesslvl.get(i).trim());
    						}
    						
    						// Check if Query return any value. If size = 0, then thr is no Read Access for the obj in SFDC.
    						strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    						QueryResult qr = con2.query(strSOQLQuery);
    						int iRecSize = qr.getSize();
    						//strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    						
    						// iRecSize = 0 if no Read access is not set. And if Read Access is not set then no Access is set 
    						Boolean bAccessLvinRTPGr8rThnAccessLvlinSFDC = true;
    						String strErrMsg = "Error: Following Access set in salesforce but not in RTP: ";
    						if (iRecSize > 0){
    							//SFDC has Read Access for Obj. RTP does not have Read Access for Obj
    					 		if (!arlistAccesslvl.contains("Read")){
    					 			bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    					 			strErrMsg = strErrMsg + "Read ";
    					 		}
    					 		
    					 		//SFDC has Create Access for Obj. RTP does not have Create Access for Obj
    					 		if (!arlistAccesslvl.contains("Create")){
    					 			strPermission = "PermissionsCreate";
    								strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    								qr = con2.query(strSOQLQuery);
    								strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    								if (strPermissionVaue.equals("true")){
    									bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    									strErrMsg = strErrMsg + "Create ";
    								}
    					 		}
    					 		
    					 		//SFDC has Edit Access for Obj. RTP does not have Edit Access for Obj
    					 		if (!arlistAccesslvl.contains("Edit")){
    					 			strPermission = "PermissionsEdit";
    								strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    								qr = con2.query(strSOQLQuery);
    								strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    								if (strPermissionVaue.equals("true")){
    									bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    									strErrMsg = strErrMsg + "Edit ";
    								}
    					 		}
    					 		
    					 		//SFDC has Delete Access for Obj. RTP does not have Delete Access for Obj
    					 		if (!arlistAccesslvl.contains("Delete")){
    					 			strPermission = "PermissionsDelete";
    								strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    								qr = con2.query(strSOQLQuery);
    								strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    								if (strPermissionVaue.equals("true")){
    									bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    									strErrMsg = strErrMsg + "Delete ";
    								}
    					 		}
    					 		
    					 		//SFDC has View All Access for Obj. RTP does not have View All Access for Obj
    					 		if (!arlistAccesslvl.contains("View All")){
    					 			strPermission = "PermissionsViewAllRecords";
    								strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    								qr = con2.query(strSOQLQuery);
    								strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    								if (strPermissionVaue.equals("true")){
    									bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    									strErrMsg = strErrMsg + "View All ";
    								}
    					 		}
    					 		
    					 		//SFDC has Modify All Access for Obj. RTP does not have Modify All Access for Obj
    					 		if (!arlistAccesslvl.contains("Modify All")){
    					 			strPermission = "PermissionsModifyAllRecords";
    								strSOQLQuery = "SELECT " +strPermission+ " FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    								qr = con2.query(strSOQLQuery);
    								strPermissionVaue = (String)qr.getRecords()[0].getField(strPermission);
    								
    								if (strPermissionVaue.equals("true")){
    									bAccessLvinRTPGr8rThnAccessLvlinSFDC = false;
    									strErrMsg = strErrMsg + "Modify All ";
    								}
    					 		}
    					 			
    						}
    						// else bAccessLvinRTPGr8rThnAccessLvlinSFDC = true
    						// if thr is no Read Access for the obj in SFDC, then no checking reqd since AccessLvl in RTP will always be Gr8r thn Access Lvl in SFDC
    						
    						boolean bOneOrMoreAccessGiven = false;
    						if (bAccessLvinRTPGr8rThnAccessLvlinSFDC){	
    							for(String strAccess : arlistAccesslvl){
    								switch(strAccess.trim()){
    									case "Read"		 : objPermission.setAllowRead(true);
	    												   bOneOrMoreAccessGiven = true;
	    												   break;
    									case "Create"	 : objPermission.setAllowCreate(true);
    												   	   bOneOrMoreAccessGiven = true;
    												   	   break;
    									case "Edit"		 : objPermission.setAllowEdit(true);
												 		   bOneOrMoreAccessGiven = true;
												 		   break;
    									case "Delete"	 : objPermission.setAllowDelete(true);
    													   bOneOrMoreAccessGiven = true;
											   			   break;
    									case "View All"	 : objPermission.setViewAllRecords(true);
    													   bOneOrMoreAccessGiven = true;
    													   break;
    									case "Modify All": objPermission.setModifyAllRecords(true);
    													   bOneOrMoreAccessGiven = true;
    													   break;
    									default			 : System.out.println("Wrong Access Level given for Object Permission: " +strAccess.trim() + " Obj: " +strObj+ " Prof: " +strProfile);
				                    		 			   FileOperations.writeToLog("Wrong Access Level given for Object Permission: " +strAccess.trim() + " Obj: " +strObj+ " Prof: " +strProfile);
				                    		 			   JOptionPane.showMessageDialog(null,"Wrong Access Level given for Object Permission: " +strAccess.trim() + " Obj: " +strObj+ " Prof: " +strProfile,"Object Access Error",JOptionPane.ERROR_MESSAGE);
    								}
    							}
    						}
    						else{
    							// no need to set Obj Access. Raise a flag
    							strErrMsg = strErrMsg + " Obj: " +strObj+ " Prof: " +strProfile;
    							System.out.println(strErrMsg);
    							FileOperations.writeToLog(strErrMsg);
    							//JOptionPane.showMessageDialog(null,strErrMsg,"Object Access Error",JOptionPane.ERROR_MESSAGE);
    						}
    						
    						if (bOneOrMoreAccessGiven)
    							alObjPermission.add(objPermission);
    						
    					} //if (!strAccessLvl.equals("BLANK")){		    	
    					else{
    						//No access level set in RTP. check if any access level is set in the app.
    						/*strSOQLQuery = "SELECT PermissionsRead FROM ObjectPermissions WHERE SobjectType = '" +strObj +"' and parentid in (select id from permissionset where PermissionSet.Profile.Name = '" +strProfile+ "')";
    						QueryResult qr = con2.query(strSOQLQuery);
    						int iRecSize = qr.getSize();
    						if (iRecSize > 0){
    							System.out.println("Error: No Access Level set in RTP but atleast Read Access level is set in salesforce");
    							FileOperations.writeToLog("Error: No Access Level set in RTP but atleast Read Access level is set in salesforce");
    							JOptionPane.showMessageDialog(null,"Error: No Access Level set in RTP but atleast Read Access level is set in salesforce","Object Access Error",JOptionPane.ERROR_MESSAGE);
    						}*/
    							
    					}
    					
    					//Assign Rec Type
    					if (!strRecTypeAss.equals("BLANK")){
    						List<String> arlistRecTypeAss = Arrays.asList(strRecTypeAss.split(","));
    						
    						for(String str : arlistRecTypeAss){
    							ProfileRecordTypeVisibility profRTV = new ProfileRecordTypeVisibility();
    							profRTV.setRecordType(strObj + "." + str.trim());
    							
    							if (str.trim().equals(strDefault)){// set that RecType in RecTypeAss column as Default which matches the rec type in the Default column in RTP
    								profRTV.setDefault(true);
    							}
    							profRTV.setVisible(true);
    							
    							alProfRecTypVis.add(profRTV);
    						}
    					}
    				}
    		    	
    			    if (alObjPermission.size() > 0 || alProfRecTypVis.size() > 0 ){
    			    	strListOfProf = strListOfProf + strProfile + "|";
        		    	Profile prof = new Profile();
        		    	prof.setFullName(strProfile);
        		    	
        		    	if (alObjPermission.size() > 0){
        		    		profObjPerms = new ProfileObjectPermissions [alObjPermission.size()];
            		    	alObjPermission.toArray(profObjPerms);
            		    	prof.setObjectPermissions(profObjPerms);
        		    	}
        		    	if (alProfRecTypVis.size() > 0){
        		    		profRecTypVis = new ProfileRecordTypeVisibility [alProfRecTypVis.size()];
            		    	alProfRecTypVis.toArray(profRecTypVis);
            			    prof.setRecordTypeVisibilities(profRecTypVis);
        		    	}
        		    	
    			    	arrayOf10Profiles.add(prof);
    			    	iCountOfProfilesNotMoreThan10 ++;
    			    	//System.out.println(strProfile);
    			    }
    		    	
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		FileOperations.writeToLog("-------------------------------------------------");
        		    	FileOperations.writeToLog("Setting Object Access for Profile: " +strListOfProf);
    			    	System.out.println("-------------------------------------------------");
    			    	System.out.println("Setting Object Access for Profile: " +strListOfProf);
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    		
    			    	updateMetadataFor10Profiles (arrayOf10Profiles, "OBJECT ACCESS");
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			else
    				FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			FileOperations.writeToLog("-------------------------------------------------");
		    	FileOperations.writeToLog("Setting Object Access for last few Profile: " +strListOfProf);
		    	System.out.println("-------------------------------------------------");
		    	System.out.println("Setting Object Access for last few Profile: " +strListOfProf);
		    	updateMetadataFor10Profiles (arrayOf10Profiles, "OBJECT ACCESS");
    		}
	    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileObjectPermissions for Profile: " +strProfile);
    		System.out.println("Error in creating ProfileObjectPermissions for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating ProfileObjectPermissions for Profile: " +strProfile,"OBJECT Access Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void createProfileLayoutAssignmentsArray () throws IOException, ConnectionException{
    	String strProfile = "";
    	try{
    		Set set = hmapProfileObjectRecTypePageLayout.entrySet();
    		Iterator iterator = set.iterator();
    		ArrayList<Profile> arrayOf10Profiles = new ArrayList<Profile> ();
    		ArrayList <ProfileLayoutAssignment> alLayoutAssignments = null;
    		ProfileLayoutAssignment[] profLayoutAss = null;
    		int iCountOfProfilesNotMoreThan10 = 0;
    		String strListOfProf = "";
    		while(iterator.hasNext()) {
    			Map.Entry mentry = (Map.Entry)iterator.next();
    			
    			strProfile = mentry.getKey().toString();
    			ArrayList <String> alObjRecTypePageLayout = new ArrayList();
    			alObjRecTypePageLayout = (ArrayList<String>) mentry.getValue();
    			
    			int iArraySize = alObjRecTypePageLayout.size();
    			if (iArraySize > 0){
    				alLayoutAssignments = new ArrayList <ProfileLayoutAssignment> ();
    				
    				int iCounter = 0;
    				for (String strFOA : alObjRecTypePageLayout){
    					String strObj = strFOA.split("#")[0].trim();
    					String strRecType = strFOA.split("#")[1].trim();
    					String strPageLayout = strFOA.split("#")[2].trim();
    					
    					int iNoOfRecType = strRecType.split(",").length;
    					List <String> alRecTyp = Arrays.asList(strRecType.split(","));
    					List <String> alPageLayout = Arrays.asList(strPageLayout.split(","));
    					
    					for(int i=0; i<iNoOfRecType; i++){
    						ProfileLayoutAssignment profLA = new ProfileLayoutAssignment();
    						
    						if (!alRecTyp.get(i).trim().equals("Master")) //this is not required for Master
    							profLA.setRecordType(strObj + "." + alRecTyp.get(i).trim());
    							
    						profLA.setLayout(strObj +"-"+ alPageLayout.get(i).trim());
    						
    						alLayoutAssignments.add(profLA);	
    					}
    				}
    		    	
    		    	if (alLayoutAssignments.size() > 0){
    		    		strListOfProf = strListOfProf + strProfile + "|";
        		    	Profile prof = new Profile();
        		    	prof.setFullName(strProfile);
        		    	
        		    	profLayoutAss = new ProfileLayoutAssignment[alLayoutAssignments.size()];
       		    		alLayoutAssignments.toArray(profLayoutAss);
        		    	prof.setLayoutAssignments(profLayoutAss);
        		    	
    		    		arrayOf10Profiles.add(prof);
    		    		iCountOfProfilesNotMoreThan10 ++;
    		    		//System.out.println(strProfile);
    		    	}
    		    	
    		    	if (iCountOfProfilesNotMoreThan10 == 10){
    		    		FileOperations.writeToLog("-------------------------------------------------");
        		    	FileOperations.writeToLog("Assigning Page Layout for Profile: " +strListOfProf);
    			    	System.out.println("-------------------------------------------------");
    			    	System.out.println("Assigning Page Layout for Profile: " +strListOfProf);
    			    	
    			    	//Profile[] profFLSArrayof10 = new Profile[arrayOf10Profiles.size()];
    			    	//arrayOf10Profiles.toArray(profFLSArrayof10);
    			    	    		
    			    	updateMetadataFor10Profiles (arrayOf10Profiles, "PAGE LAYOUT ASSIGNMENT");
    			    	
    			    	arrayOf10Profiles = new ArrayList<Profile> ();
    			    	iCountOfProfilesNotMoreThan10 = 0;
    			    	strListOfProf = "";
    		    	}
    			}
    			else
    				FileOperations.writeToLog("Nothing To Set for Profile: " +strProfile);
    		}
    		if (iCountOfProfilesNotMoreThan10 > 0){
    			FileOperations.writeToLog("-------------------------------------------------");
		    	FileOperations.writeToLog("Assigning Page Layout for last few Profile: " +strListOfProf);
		    	System.out.println("-------------------------------------------------");
		    	System.out.println("Assigning Page Layout for last few Profile: " +strListOfProf);
		    	updateMetadataFor10Profiles (arrayOf10Profiles, "PAGE LAYOUT ASSIGNMENT");
    		}
	    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating ProfileLayoutAssignment for Profile: " +strProfile);
    		System.out.println("Error in creating ProfileLayoutAssignment for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating ProfileLayoutAssignment for Profile: " +strProfile,"PAGE LAYOUT ASSIGNMNET Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void createPermissionSetFieldLvlSecurityArray(String strProfile, ArrayList<String> alFieldObjectAccess) throws IOException, ConnectionException{
    	try{
	    	int iArraySize = alFieldObjectAccess.size();
	    	PermissionSetFieldPermissions[] fieldPermissions = new PermissionSetFieldPermissions[iArraySize];
	    	
	    	int iCounter = 0;
	    	for (String strFOA : alFieldObjectAccess){
				String strFieldObj = strFOA.split("#")[0];
				String strAccess = strFOA.split("#")[1];
				
				fieldPermissions[iCounter] = new PermissionSetFieldPermissions();
		    	fieldPermissions[iCounter].setField(strFieldObj);
		    	
		    	//Read ONLY: setEditable(false).. Read/Write: setEditable(true)
		    	if (strAccess.equalsIgnoreCase("Read ONLY"))
		    		fieldPermissions[iCounter].setEditable(false); 
		    	else if (strAccess.equalsIgnoreCase("Read/Write"))
		    		fieldPermissions[iCounter].setEditable(true); 
		    	else
		    		FileOperations.writeToLog("Error: Access Level provided in excel - " +strAccess + "for Permission Set: " +strProfile);
		    	
		    	fieldPermissions[iCounter].setReadable(true);
		    	iCounter ++;
			}
	    	
	    	FileOperations.writeToLog("-------------------------------------------------");
	    	FileOperations.writeToLog("Setting FLS for Permission Set: " +strProfile);	 
	    	
	    	if (iArraySize > 0)
	    		setFLSforPermissionSet (strProfile, fieldPermissions);
	    	else
	    		FileOperations.writeToLog("Nothing To Set");
	    		
    	}catch(Exception e){
    		FileOperations.writeToLog("Error in creating PermissionSetFieldPermissions for Profile: " +strProfile);
    		System.out.println("Error in creating PermissionSetFieldPermissions for Profile: " +strProfile);
    		e.printStackTrace();
    		JOptionPane.showMessageDialog(null,"Error in creating PermissionSetFieldPermissions for Profile: " +strProfile,"PERMISSION SET ASSIGNMENT Error",JOptionPane.ERROR_MESSAGE);
    	}
    }
    
    public static void setFLSforPermissionSet(String strPermissionSet, PermissionSetFieldPermissions[] fieldPermissions) throws ConnectionException, IOException{
    	try{
    		PermissionSet permSet = new PermissionSet();	
	    	
    		permSet.setFullName(strPermissionSet);
    		permSet.setFieldPermissions(fieldPermissions);
	    	
	    	SaveResult[] arsTab =  con.updateMetadata(new Metadata[] {permSet});
	    	
	    	for (SaveResult r : arsTab) {
	            if (r.isSuccess()) {
	                System.out.println("Updated component: " + r.getFullName());
	                FileOperations.writeToLog("Updated component: " + r.getFullName());
	            } else {
	                System.out.println("Errors were encountered while updating " +r.getFullName());
	                for (com.sforce.soap.metadata.Error e : r.getErrors()) {
	                    System.out.println("Error message: " + e.getMessage());
	                    FileOperations.writeToLog("Error message: " + e.getMessage());
	                    JOptionPane.showMessageDialog(null,e.getMessage(),"FLS Update Error",JOptionPane.ERROR_MESSAGE);
	                    System.out.println("Status code: " + e.getStatusCode());
	                    FileOperations.writeToLog("Status code: " + e.getStatusCode());
	                    JOptionPane.showMessageDialog(null,e.getStatusCode(),"FLS Update Error",JOptionPane.ERROR_MESSAGE);
	                }
	            }
	        }
    	}catch(ConnectionException ce){
    		ce.printStackTrace();
    	}
    }

    public static void updateMetadataFor10Profiles (ArrayList<Profile> arrayOf10Profiles, String strMetadataToUpdate) throws ConnectionException, IOException{
    	try{
    		Metadata[] meta = new Metadata[arrayOf10Profiles.size()];
	    	for(int i=0; i<arrayOf10Profiles.size(); i++)
	    		meta[i] = arrayOf10Profiles.get(i);
	    	
	    	SaveResult[] arsTab =  con.updateMetadata(meta);
	    	
	    	//SaveResult[] arsTab =  con.updateMetadata(new Metadata[] {arrayOf10Profiles.get(0)});
	    	
	    	for (SaveResult r : arsTab) {
	            if (r.isSuccess()) {
	                System.out.println("Updated component: " + r.getFullName());
	                FileOperations.writeToLog("Updated component: " + r.getFullName());
	            } else {
	                System.out.println("Errors were encountered while updating " +r.getFullName());
	                for (com.sforce.soap.metadata.Error e : r.getErrors()) {
	                    System.out.println("Error message: " + e.getMessage());
	                    FileOperations.writeToLog("Error message: " + e.getMessage());
	                    JOptionPane.showMessageDialog(null,e.getMessage(), strMetadataToUpdate +"Error",JOptionPane.ERROR_MESSAGE);
	                    System.out.println("Status code: " + e.getStatusCode());
	                    FileOperations.writeToLog("Status code: " + e.getStatusCode());
	                    JOptionPane.showMessageDialog(null,e.getStatusCode(), strMetadataToUpdate +"Error",JOptionPane.ERROR_MESSAGE);
	                }
	            }
	        }
    	}catch(ConnectionException ce){
    		ce.printStackTrace();
    	}
    }

}
