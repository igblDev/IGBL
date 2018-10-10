
	wFirstName    = objsName.Fields("firstName").Value
	wLastName     = objsName.Fields("lastName").Value

	objsBarps.Open "SELECT barps FROM tbl_barps t WHERE t.first = '" & wFirstName & "'  and t.last ='"&wLastName & "' ", objConn,1,1					
	
	objRSAll.Open "SELECT avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
								"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP " &_
								"FROM tblLast5 t " & _
								"WHERE t.first = '" & wFirstName & "'  and t.last ='"&wLastName & "' ", objConn,1,1

	avgMP    = objRSAll.Fields("avgMP").Value
	avgBlks  = objRSAll.Fields("avgBlks").Value
	avgAst   = objRSAll.Fields("avgAst").Value
	avgReb   = objRSAll.Fields("avgReb").Value
	avgPts   = objRSAll.Fields("avgPts").Value
	avgStl   = objRSAll.Fields("avgStl").Value
	avg3pt   = objRSAll.Fields("avg3pt").Value
	avgTo    = objRSAll.Fields("avgTo").Value
	avgBarps = objRSAll.Fields("avgBarps").Value
	
