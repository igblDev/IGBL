	<%	
	objParams.Open  "SELECT * FROM tblParameterCtl WHERE param_name = 'PLAYOFF_START_DATE' ", objConn,3,3,1
	dQtrDate = objParams.Fields("param_date").value
	objParams.Close
	
	objParams.Open  "SELECT * FROM tblParameterCtl WHERE param_name = 'PO_SEMIS' ", objConn,3,3,1
	dSemiDate = objParams.Fields("param_date").value
	objParams.Close
	
	'Response.Write "PLAYOFF_START_DATE Date  = "&dQtrDate&"<br>"
	'Response.Write "PO_SEMIS Date = "&dSemiDate&"<br>"
	
		
	objrsATL.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.ATL = 1 and tblGameGrid.gamedate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsBKN.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.BKN = 1 and tblGameGrid.gamedate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsBOS.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.BOS = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsCHA.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.CHA = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsCHI.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.CHI = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsCLE.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.CLE = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsDAL.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.DAL = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsDEN.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.DEN = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsDET.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.DET = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsGSW.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.GSW = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsHOU.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.HOU = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsIND.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.IND = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsLAC.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.LAC = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsLAL.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.LAL = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1	
	objrsMEM.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.MEM = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsMIA.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.MIA = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsMIL.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.MIL = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsMIN.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.MIN = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsNOP.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.NOP = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsNYK.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.NYK = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsOKC.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.OKC = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsORL.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.ORL = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsPHI.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.PHI = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsPHX.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.PHX = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsPOR.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.POR = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsSAC.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.SAC = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsSAS.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.SAS = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsTOR.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.TOR = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsUTA.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.UTA = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
	objrsWAS.Open		  "SELECT * FROM tblGameGrid WHERE tblGameGrid.WAS = 1 and tblGameGrid.GameDate >= cdate('"&dQtrDate&"') and tblGameGrid.gamedate < cdate('"&dSemiDate&"') ", objConn,3,3,1
%>