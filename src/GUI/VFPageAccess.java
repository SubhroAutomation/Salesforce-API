package GUI;

import static GUI.PostDepTool.hmapProfileVFPageAccess;
import static GUI.PostDepTool.alVFPageProfilesfrmRTP;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadVFPagefromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

public class VFPageAccess {
	public static void setVFPageAccess() throws Exception{
		FileOperations.createFile("VFPage");	
		
		int iStartRow = -1;
		iStartRow = ReadVFPagefromRTPExcel.getStartRowIndexAfterAllProfTaginRTP_VFPage();
		
		alVFPageProfilesfrmRTP = ReadVFPagefromRTPExcel.getListofVFPageProfilesfrmRTP (iStartRow);
		ReadVFPagefromRTPExcel.getCombinationofAllVFPageProfinRTP(iStartRow);
		
		int iStartRowAllProf = -1;
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "All Profiles-Start", 0, 0) + 1;
		ReadVFPagefromRTPExcel.getCombinationofVFPagefrmALLPROFILES (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();
		
		MetaDataAPI.createProfileApexPageAccessArray();
	}
}
