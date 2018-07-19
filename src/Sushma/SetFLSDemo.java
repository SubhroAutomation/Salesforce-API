package Sushma;

import com.sforce.soap.enterprise.EnterpriseConnection;
import com.sforce.soap.enterprise.LoginResult;
import com.sforce.soap.metadata.AsyncResult;
import com.sforce.soap.metadata.DescribeMetadataObject;
import com.sforce.soap.metadata.DescribeMetadataResult;
import com.sforce.soap.metadata.Metadata;
import com.sforce.soap.metadata.MetadataConnection;
import com.sforce.soap.metadata.Profile;
import com.sforce.soap.metadata.ProfileApexClassAccess;
import com.sforce.soap.metadata.ProfileApexPageAccess;
import com.sforce.soap.metadata.ProfileFieldLevelSecurity;
import com.sforce.soap.metadata.ProfileObjectPermissions;
import com.sforce.soap.metadata.ProfileRecordTypeVisibility;
import com.sforce.soap.metadata.SaveResult;
import com.sforce.soap.metadata.UpdateMetadata_element;
import com.sforce.ws.ConnectionException;
import com.sforce.ws.ConnectorConfig;
import com.sun.xml.internal.ws.util.MetadataUtil;

import Utilities.ExcelPOI;
import Utilities.FileOperations;

public class SetFLSDemo {
	public static MetadataConnection con = null;
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		login();
		setFLSforProfile();
	}
	
	public static MetadataConnection login() throws ConnectionException {
        final String USERNAME = "sysuser@sunpower.com.l3accteam"; //"subhra.bikashdas@sunpowercorp.com.testdeploy";
        final String PASSWORD = "Solar123"; //925i5XnTrtXmrVW2seIoeIke";
        final String URL =  "https://test.salesforce.com/services/Soap/c/29.0";
       
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
        return (new EnterpriseConnection(config)).login(username, password); 
    }
   
    public static void setFLSforProfile() throws ConnectionException{
    	try{
    		/*	•	Global Sales Ops
				•	System Administrator – Integration
				•	Company Executive
				•	System Administrator - No Customization
				•	System Administrator - No DDL
				•	System Administrator - No user Admin
			*/
    		
    		String strProfile = "System Administrator - No user Admin";
    		//String strProfile = "AU Partner Installer";
	    	Profile prof = new Profile();
	    	
	    	ProfileFieldLevelSecurity[] fieldPermissions = new ProfileFieldLevelSecurity[1];
	    	fieldPermissions[0] = new ProfileFieldLevelSecurity();
	    	fieldPermissions[0].setField("NH_Community__c.Account_Manager__c");
	    	fieldPermissions[0].setEditable(false);
	    	fieldPermissions[0].setReadable(true);
	    	
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

}
