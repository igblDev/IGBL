
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
			<span style="font-size:12px;color:black;" class="pull-right"><strong>Sortable Season Stats</strong></span>
			</div>
		</div>
		</br>
		<div class="row">
			<div class="col-md-12 col-sm-12 col-xs-12">
				<table class="table table-striped table-bordered table-custom-black table-condensed" width="100%" class="display" id="example2">
				<thead>
					<tr>
						<th class="big" style="vertical-align:middle;text-align:center;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i><span></th>
						<th class="big">Player</th>
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
						<th class="big hidden-xs" style="text-decoration: underline;">3D</th>
					</tr>
				</thead>
				<tbody>
					<%
						While Not objRSSort.EOF					 
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
						<td class="big" style="vertical-align: middle;font-weight:bold;"><%=round(objRSSort.Fields("barps").Value,2)%></td>							
						<%if objRSSort.Fields("statusDesc").Value  = "On Team"then	%>
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("FirstName").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("LastName").Value,12)%></a>
									<span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
							</td>
						<%else%>							
							<td class="big" style="vertical-align:middle;text-align:left;">
								<a class="blue big" href="playerprofile.asp?pid=<%=objRSSort.Fields("PID").Value %>"><%=left(objRSSort.Fields("FirstName").Value,1)%>.&nbsp;<%=left(objRSSort.Fields("LastName").Value,12)%></a>&nbsp;<small style="color:red;font-weight:bold;">[FA]</small>
									<span class="greenTrade text-uppercase"><%=objRSSort.Fields("teamShortName").Value %></span>&nbsp;<span class="orange"><%=objRSSort.Fields("pos").Value %></span>
							</td>								
						<%end if%>
							<td class="big" style="vertical-align: middle;"><%=objRSSort.Fields("gp").Value%></td>
							<td class="big hidden-xs" style="vertical-align: middle;"><%=objRSSort.Fields("min").Value%></td>
						  <td class="big hidden-xs" style="vertical-align: middle;"><%=objRSSort.Fields("usage").Value%></td>		
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