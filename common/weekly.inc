<%
	
	Dim hotStartDate,hotEndDate,objRSSort,objRSPID
	
	hotEndDate   = date()
	hotStartDate = (date() - 10)
	
	Set objRSSort    = Server.CreateObject("ADODB.RecordSet")
	Set objRSPID     = Server.CreateObject("ADODB.RecordSet")
	Set objRSDate    = Server.CreateObject("ADODB.RecordSet")

	objRSDate.Open  "SELECT MAX(gameDate) as EndDate, " &_ 
					        "MAX(GameDate) - 6 as StartDate, " &_
					        "MIN(GameDate) as FirstGame " & _
 	                "FROM tblLast5 ", objConn,3,3,1
	
	if IsNull(objRSDate.Fields("EndDate")) then		
		hotEndDate = date()
		hotStartDate = date()
	else
		hotEndDate = objRSDate.Fields("EndDate").Value				
		if objRSDate.Fields("FirstGame").Value > objRSDate.Fields("StartDate").Value then
			hotStartDate = objRSDate.Fields("FirstGame").Value
		else
		    hotStartDate = objRSDate.Fields("StartDate").Value
		end if		
	end if	
	
	
	objRSDate.Close								
%>
	</br>
	<div class="row">	
		<div class="col-md-12 col-sm-12 col-xs-12">	
			<span style="font-size:12px;color:red;" class="pull-right"><strong><span style="color:black;"><%= hotStartDate %></span> to <span style="color:black;"><%= hotEndDate %></strong></span>
		</div>
	</div>
	</br>

<%
	objRSSort.Open    "SELECT  first, last, count(1) as Games, " &_
                    "avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
										"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP, avg(usage) as avgUSG " &_
                    "FROM tblLast5 t " &_
                    "WHERE gamedate >= #"&hotStartDate&"# " &_
                    "GROUP by first, last " &_
					          "ORDER by avg(BARPTot) desc ", objConn,3,3,1
%>	




		<table class="table table-bordered table-custom-black table-condensed" width="100%" class="display" id="weeks1">
		<thead>
			<tr>
			<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
			<th class="big">Player</th>
			<th class="big style="text-decoration: underline;">G</th>
			<th class="big hidden-xs" style="text-decoration: underline;">M</th>
			<th class="big hidden-xs" style="text-decoration: underline;">U</th>
			<th class="big" style="text-decoration: underline;">B</th>
			<th class="big" style="text-decoration: underline;">A</th>
			<th class="big" style="text-decoration: underline;">R</th>
			<th class="big" style="text-decoration: underline;">P</th>
			<th class="big" style="text-decoration: underline;">S</th>
			<th class="big" style="text-decoration: underline;">3</th>
			<th class="big" style="text-decoration: underline;">T</th>
		</tr>
		</thead>
		<tbody>
			<%
				While Not objRSSort.EOF		
				wFirstName = objRSSort.Fields("first").Value
				wLastName  = objRSSort.Fields("last").Value
										
				objRSPID.Open "SELECT p.pid, p.firstName, p.lastName,p.POS,p.ownerID, n.teamShortName " & _
											"FROM tblPlayers p,  tblNBATEAMS n " & _
											"where  p.NBATeamID = n.NBATID " & _
											"and p.firstName = '"&wFirstName&"'  " & _
											"and p.lastName = '"&wLastName&"' ", objConn,1,1
				
				if objRSPID.RecordCount > 0 then 
				   wPID      = objRSPID.Fields("PID").Value 
				   wTeam     = objRSPID.Fields("teamShortName").Value
				   wPOS      = objRSPID.Fields("pos").Value
				   wPOwnerID = objRSPID.Fields("ownerID").Value
				else
				   wPID      = 9994
				   wTeam     = "NA"
				   wPOS      = "NA"
				   wPOwnerID = 0
				end if
				objRSPID.Close

			%>
					<%if wPOwnerID > 0 then%>
						<%if wPOwnerID = ownerid then%>
							<tr class="success">	
						<%else%>
							<tr style="background-color:white;">		
						<%end if%>	
					<%else%>
						<tr class="warning"> 
					<%end if%>		
					<td class="big" style="vertical-align: middle;font-weight:bold;"><%=round(objRSSort.Fields("avgBarps").Value,2)%></td>							
					<td class="big" style="vertical-align: middle;text-align:left;">
					<%if wPOwnerID > 0 then%>					
						<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSSort.Fields("first").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("last").Value,15)%></a>&nbsp;<span class="greenTrade text-uppercase"><%=wTeam%></span>&nbsp;<span class="orange text-uppercase"><%=wPos%></span>
					<%else%>
						<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSSort.Fields("first").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("last").Value,15)%></a>&nbsp;<small style="color:red;font-weight:bold;">[FA]</small>&nbsp;<span class="greenTrade text-uppercase"><%=wTeam%></span>&nbsp;<span class="orange text-uppercase"><%=wPos%></span>
					<%end if%>
					
					
					</td>
					<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("Games").Value%></td>
					<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgMP").Value,0)%></td>
					<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgUSG").Value,0)%></td>									
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgBlks").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgAst").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgReb").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgPts").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgStl").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avg3pt").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("avgTo").Value,0)%></td>
				</tr>		
				<%
					objRSSort.MoveNext
					Wend
				%>
		</tbody>	
		</table>
		<br>

<%
objRSNBASked.Close
objRSPID.Close
ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
