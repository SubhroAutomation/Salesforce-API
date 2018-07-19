package GUI;

import static GUI.PostDepTool.alApexClassProfilesfrmRTP;
import static GUI.PostDepTool.hmapProfileApexClassAccess;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadApexClassfromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

public class ApexClassAccess {
	public static void setApexClassAccess() throws Exception{
		FileOperations.createFile("ApexClass");	
		
		int iStartRow = -1;
		iStartRow = ReadApexClassfromRTPExcel.getStartRowIndexAfterAllProfTaginRTP_ApexClass();
		
		alApexClassProfilesfrmRTP = ReadApexClassfromRTPExcel.getListofApexClassProfilesfrmRTP (iStartRow);
		ReadApexClassfromRTPExcel.getCombinationofAllApexClassProfinRTP(iStartRow);
		
		int iStartRowAllProf = -1;
		int iStartRowofClassAccess = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Class Access:", 0, 0);
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "All Profiles-Start", 0, iStartRowofClassAccess) + 1;
		ReadApexClassfromRTPExcel.getCombinationofApexClassfrmALLPROFILES (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();
		
		MetaDataAPI.createProfileApexClassAccessArray();
	}
}
