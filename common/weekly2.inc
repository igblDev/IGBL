<%
	
	Dim hotStartDate,hotEndDate,objRSSort2,objRSPID2
	
	hotEndDate   = date()
	hotStartDate = (date() - 10)
	
	Set objRSSort2    = Server.CreateObject("ADODB.RecordSet")
	Set objRSPID2     = Server.CreateObject("ADODB.RecordSet")
	Set objRSDate     = Server.CreateObject("ADODB.RecordSet")

	objRSDate.Open  "SELECT MAX(gameDate) as EndDate, " &_ 
									"MAX(GameDate)-(select param_amount-1 from tblParameterCtl where param_name = 'WHOS_HOT_NBR_DAYS') as StartDate, " &_
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
			<span style="font-size:12px;color:black;" class="pull-right"><strong>Sortable Stats</strong></span></br>		
			<span style="font-size:12px;color:red;" class="pull-right"><strong><span style="color:black;"><%= hotStartDate %></span> to <span style="color:black;"><%= hotEndDate %></strong></span>
		</div>
	</div>
	</br>

<%
	objRSSort2.Open   "SELECT  first, last, count(*) as Games, " &_
                    "avg(x3p) as avg3pt, avg(TRB) as avgReb, avg(AST) as avgAst, avg(STL) as avgStl, avg(BLK) as avgBlks, " &_
										"avg(TOV) as avgTo, avg(PTS) as avgPts, avg(BARPTot) as avgBarps, avg(MP) as avgMP, avg(usage) as avgUSG " &_
                    "FROM tblLast5 t " &_
                    "WHERE gamedate = date() - 13 " &_
                    "GROUP by first, last " &_
										"ORDER by avg(BARPTot) desc ", objConn,3,3,1
%>	




		<table class="table table-bordered table-custom-black table-condensed" width="100%" class="display" id="weeks2">
		<thead>
			<tr>
				<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
				<th class="big" style="text-decoration: underline;">Name</th>
				<th class="big" style="text-decoration: underline;">G</th>
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
				While Not objRSSort2.EOF		
				wFirstName = objRSSort2.Fields("first").Value
				wLastName  = objRSSort2.Fields("last").Value
										
				objRSPID2.Open "SELECT p.pid, p.firstName, p.lastName,p.POS, n.teamShortName " & _
											"FROM tblPlayers p,  tblNBATEAMS n " & _
											"where  p.NBATeamID = n.NBATID " & _
											"and p.firstName = '"&wFirstName&"'  " & _
											"and p.lastName = '"&wLastName&"' ", objConn,1,1
				
				wPID      = objRSPID2.Fields("PID").Value 
				wTeam     = objRSPID2.Fields("teamShortName").Value
				wPOS      = objRSPID2.Fields("pos").Value
				objRSPID2.Close

			%>
				<tr style="background-color:white;">
					<td class="big" style="vertical-align: middle;font-weight:bold;"><%=round(objRSSort2.Fields("avgBarps").Value,2)%></td>							
					<td class="big" style="vertical-align: middle;text-align:left;">								
						<a class="blue" href="playerprofile.asp?pid=<%=wPID %>"><%=left(objRSSort2.Fields("first").Value,1)%>.&nbsp;<%=left(objRSSort2.Fields("last").Value,15)%></a>&nbsp;<span class="greenTrade text-uppercase"><%=wTeam%></span>&nbsp;<span class="orange text-uppercase"><%=wPos%></span>
					</td>
					<td class="big" style="vertical-align: middle;"><%=objRSSort2.Fields("Games").Value%></td>
					<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgMP").Value,0)%></td>
					<td class="big hidden-xs" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgUSG").Value,0)%></td>									
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgBlks").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgAst").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgReb").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgPts").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgStl").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avg3pt").Value,0)%></td>
					<td class="big" style="vertical-align: middle;"><%=round(objRSSort2.Fields("avgTo").Value,0)%></td>	
				</tr>		
				<%
					objRSSort2.MoveNext
					Wend
				%>
		</tbody>	
		</table>
		<br>

<%
objRSSort2.Close
objRSPID2.Close
ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
