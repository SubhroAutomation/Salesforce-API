package GUI;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import Components.MetaDataAPI;
import Components.ReadPageLayoutfromRTPExcel;
import Utilities.ExcelPOI;
import Utilities.FileOperations;

import static GUI.PostDepTool.hmapProfileObjectRecTypePageLayout;
import static GUI.PostDepTool.alPageLayoutProfilesfrmRTP;

public class PageLayoutAssignment {
	public static void assignPageLayout() throws Exception{
		FileOperations.createFile("PageLayout");	
		
		int iStartRow = -1;
		iStartRow = ReadPageLayoutfromRTPExcel.getStartRowIndexAfterAllProfTaginRTP_PageLayout();
		
		alPageLayoutProfilesfrmRTP = ReadPageLayoutfromRTPExcel.getListofPageLayoutProfilesfrmRTP (iStartRow);
		ReadPageLayoutfromRTPExcel.getCombinationofAllObjAccessRecTypAssDefaultProfinRTP (iStartRow);
		
		int iStartRowAllProf = -1;
		int iStartRowofPageLayout = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "Page Layout:", 0, 0);
		iStartRowAllProf = ExcelPOI.GetRowIndexofValueinCol("VFPage,Class,Obj,PageLayout", "All Profiles-Start", 0, iStartRowofPageLayout) + 1;
		ReadPageLayoutfromRTPExcel.getCombinationofObjAccessRecTypAssDefaultfrmALLPROFILES (iStartRowAllProf);
		
		//RTPExcel.createProfileFieldObjAccessSheet ();
		
		MetaDataAPI.createProfileLayoutAssignmentsArray();
	}
}
