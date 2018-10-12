
<%
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	Set objRSSort    = Server.CreateObject("ADODB.RecordSet")
	Set objRSNBASked = Server.CreateObject("ADODB.RecordSet")

%>
</br>
<%
	objRSSort.Open  "SELECT * FROM qry_PlayerAll order by barps desc", objConn,3,3,1
%>
<!--#include virtual="Common/headerMain.inc"-->

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" method="POST" language="JavaScript" name="FrontPage_Form1">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example">
				<thead>
					<tr>
						<th class="big">Player</th>
						<th class="big" style="text-decoration: underline;">A</th>
						<th class="big" style="text-decoration: underline;">5</th>
						<th class="big" style="text-decoration: underline;">B</th>
						<th class="big" style="text-decoration: underline;">A</th>
						<th class="big" style="text-decoration: underline;">R</th>
						<th class="big" style="text-decoration: underline;">P</th>
						<th class="big" style="text-decoration: underline;">S</th>
						<th class="big" style="text-decoration: underline;">3</th>
						<th class="big" style="text-decoration: underline;">T</th>
						<th class="big hidden-xs" style="text-decoration: underline;">D</th>
					</tr>
				</thead>
				<tbody>
					<%
					While Not objRSSort.EOF
					 
					 'Response.Write "DO I PLAY TODAY  " &objRSSort.Fields("NBATeamID").Value& "<br> "
					 'Response.Write "PLAYER LAST NAME  " &objRSSort.Fields("lastName").Value& "<br> "
					 
					 objRSNBASked.Open "Select * from NBAINDTMSKed where NBAINDTMSKed.GameDay = date() and NBAINDTMSKed.NBATeam = "&objRSSort.Fields("NBATeamID").Value, objConn,3,3,1
					 'Must assign it to a variable because the code give inconsitent values when you evaluate the Time field when no rows are returned.
					 if objRSNBASked.RecordCount > 0 then
					     'Response.Write "RECORD COUNT " &objRSNBASked.RecordCount& "<br> "
							 wTipTime = objRSNBASked.Fields("GameTime").Value
					 else
							 wTipTime = "12:00:00 AM"
					 end if						   
					 if len(objRSNBASked.Fields("GameTime").Value) = 10 then
							wtime = Left(objRSNBASked.Fields("GameTime").Value,4) & Right(objRSNBASked.Fields("GameTime").Value,3)
					 else
							wtime = Left(objRSNBASked.Fields("GameTime").Value,5) & Right(objRSNBASked.Fields("GameTime").Value,3)
					 end if	
					 
					
					%>
					
					<%if objRSSort.Fields("ownerID").Value > 0 then%>
						<%if objRSSort.Fields("ownerID").Value = ownerid then%>
							<tr class="success">	
						<%else%>
							<tr style="background-color:white;">		
						<%end if%>	
					<%else%>
						<tr class="warning"> 
					<%end if%>	
						<%if objRSSort.Fields("statusDesc").Value  = "On Team"then	%>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("FirstName").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("LastName").Value,12)%></a>
								<small>
									<!--</br><span class="redTrade  text-uppercase"><%=objRSSort.Fields("shortName").Value %></span>-->
									<small><span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span></small>
								</small>
							</td>
						<%else%>
								<% if objRSNBASked.RecordCount > 0 then %>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue big" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("FirstName").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("LastName").Value,12)%></a>
								<small>
									<span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
								</small>
							</td>
								<%else%>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue big" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("LastName").Value,12)%></a>
								<small>
									</br><span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
								</small>
							</td>
								<%end if%>		
						<%end if%>
							<td class="big" style="vertical-align: middle;"><%=round(objRSSort.Fields("barps").Value,0)%></td>		
							<% if CInt(objRSSort.Fields("l5barps").Value) > CInt(objRSSort.Fields("barps").Value) then %>
								<td class="big" style="vertical-align:middle;text-align:center;background-color:#468847;color:white;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
							<% elseif CInt(objRSSort.Fields("barps").Value) > CInt(objRSSort.Fields("l5barps").Value) then%>
								<td class="big" style="vertical-align:middle;text-align:center;background-color:#9a1400;color:white;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
							<%else%>
								<td class="big" style="vertical-align:middle;text-align:center;background-color:gold;color:black;"><%=round(objRSSort.Fields("l5barps").Value,0)%></td>									
							<%end if %>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("blk").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("ast").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("reb").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("ppg").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("stl").Value%></td>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("three").Value%></td>
							<td class="big"  style="vertical-align: middle;"><%=objRSSort.Fields("to").Value%></td>
							<td class="big hidden-xs" style="vertical-align: middle;"><%=objRSSort.Fields("numTdbls").Value%></td>
						</tr>		
						<%
						objRSNBASked.Close
						objRSSort.MoveNext
						
						Wend
						%>
				</tbody>	
				</table>
				<br>
			</div>
		</div>
</form>
<%
objRSNBASked.Close
ObjConn.Close
Set objConn = Nothing
Session.CodePage = Session("FP_OldCodePage")
Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>