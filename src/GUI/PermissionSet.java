package GUI;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadPermissionSetfromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

import static GUI.PostDepTool.hmapPermissionSetFieldObjAccess;
import static GUI.PostDepTool.alPermissionSetsfrmRTP;

public class PermissionSet {
	public static void setPermissionSet() throws Exception{
		FileOperations.createFile("PermissionSet");	
		
		int iStartRow = -1;
		iStartRow = ReadPermissionSetfromRTPExcel.getStartRowIndexAfterAllPermissionTaginRTP();
		
		alPermissionSetsfrmRTP = ReadPermissionSetfromRTPExcel.getListofPermissionSetsfrmRTP (iStartRow);
		ReadPermissionSetfromRTPExcel.getCombinationofAllFieldObjAccessPermSetinRTP (iStartRow);
		
		int iStartRowAllProf = -1;
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("Permission Set changes", "ALL Permission Set-Start", 0, 0) + 1;
		ReadPermissionSetfromRTPExcel.getCombinationofFieldObjAccfrmALLPERMISSIONSETS (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();

		
		/*Set set = hmapPermissionSetFieldObjAccess.entrySet();
		Iterator iterator = set.iterator();
		while(iterator.hasNext()) {
			Map.Entry mentry = (Map.Entry)iterator.next();
			System.out.print("PERMISSION SET: "+ mentry.getKey() + ".... COMBINATIONS: ");
			System.out.println(mentry.getValue());
			FileOperations.writeToLog("-------------------------------------------------");
			FileOperations.writeToLog("PermissionSet: "+ mentry.getKey());
			FileOperations.writeToLog("Field/Obj/Access Entries: "+ mentry.getValue().toString());
		}*/
		
		Set set = hmapPermissionSetFieldObjAccess.entrySet();
		Iterator iterator = set.iterator();
		while(iterator.hasNext()) {
			Map.Entry mentry = (Map.Entry)iterator.next();
			
			String strPermissionSet = mentry.getKey().toString();
			ArrayList <String> alPermissionSetsfrmRTP = new ArrayList();
			alPermissionSetsfrmRTP = (ArrayList<String>) mentry.getValue();
			if (alPermissionSetsfrmRTP.size() > 0)
				MetaDataAPI.createPermissionSetFieldLvlSecurityArray (strPermissionSet, alPermissionSetsfrmRTP);
		}
	}
}
