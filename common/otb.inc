<%
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"

	GetAnyParameter "Action", sAction
	GetAnyParameter "var_teamname", teamname

	Set objRSteams     = Server.CreateObject("ADODB.RecordSet")
	Set objRSteamsneeds= Server.CreateObject("ADODB.RecordSet")
	Set objRSontheBlock= Server.CreateObject("ADODB.RecordSet")
	Set objRSBlock     = Server.CreateObject("ADODB.RecordSet")

	objRSteams.Open "SELECT * FROM qryowners where ownerId <> 99  order by ownerId", objConn,3,3,1

	dim count
 	count = CInt(count )
 	count = 1
%>
</br>
<form action="dashboard.asp" name="frmotb" method="POST">
	<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
	<input type="hidden" name="var_teamname" value="<%= teamname %>" />

	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
					<table class="table table-bordered table-custom-black table-responsive table-condensed">
						<tr style="background-color: #ececec;color:black">
							<th class ="big" style="color:black;width:40%">Offered</th>
							<th class="big" style="color:black;text-align:center;width:10%">Needs</th>
							<th class="big" style="color:black;width:40%">Comments</th>
						</tr>
						<%
						 While Not objRSteams.EOF
						%>							
						<%if objRSteams.Fields("OwnerID").value = ownerid then%>
						<tr style="background-color:white;">
							<th class="big" colspan="3"><span style="font-weight:bold;"><%= objRSteams.Fields("teamname").Value%></span></th>
						</tr>	
						<%elseif myTradeInd = false or objRSteams.Fields("acceptTradeOffers").Value = false then%>
						<tr style="background-color: yellow">
							<td colspan="3" class="text-uppercase"  style="vertical-align: sub;text-align:center;" ><i class="fas fa-ban red fa-lg" style="vertical-align: sub;text-align:center;"></i>&nbsp;<span style="font-weight:bold;background-color: yellow;vertical-align: middle;"><%= objRSteams.Fields("teamname").Value%> TRADE DESK IS CLOSED!</span>&nbsp;<i class="fas fa-ban red fa-lg" style="vertical-align: sub;"></i></td>
						</tr>	
						<%else%>
						<tr>
							<td class="big" style="text-transform: uppercase;vertical-align:middle;" colspan="3">
								<button class="btn btn-link  btn-block btn-trades " value="Continue;<%= Trim(objRSteams.Fields("OwnerID").value) %>;<%= objRSteams.Fields("TeamName").value%>" name="Action" type="submit">Trade w/<%= objRSteams.Fields("TeamName").value%>&nbsp;<small>(<%= objRSteams.Fields("ShortNAme").Value%>)</small>
							</td>
						</tr>
						<%end if %>
						<tr>
						<%
							objRSBlock.Open "SELECT * FROM tblPlayers p, tbl_barps b, tblNBATeams t WHERE (p.OntheBlock = True AND p.ownerID = " & count & ") and p.firstname = b.first and p.lastname = b.last and p.nbateamid = t.nbatid order by b.barps desc ", objConn,3,3,1
						%>
							<td class="big">
								<table class="table table-responsive table-condensed">
									<%
									While Not objRSBlock.EOF
									%>
									<tr>
										<td class="big" style="width:33%;background-color:white;text-align:left;"><a class="blue" href="playerprofile.asp?pid=<%=objRSBlock.Fields("pid").Value %>" target="_self"><%= left(objRSBlock.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSBlock.Fields("lastName").Value,10)%></a>&nbsp;<small><span class="greenTrade"><%=objRSBlock.Fields("TeamshortName").Value%></span>&nbsp;<span class="orange"><%=objRSBlock.Fields("POS").Value%></span></small></td>	
									</tr>
									<%
										objRSBlock.MoveNext
										Wend	
									%>
								</table>
							</td>
							<%
								objRSBlock.Close
								objRSteamsneeds.Open "SELECT ontheblockneedscen, ontheblockneedsfor,ontheblockneedsgua,ontheblockneedsfc, ontheblockneedsgf FROM tblowners WHERE ownerid = " & count & " ", objConn
							%>
							<td class="big" style="vertical-align:text-top;display:table-cell;">
							<% if objRSteamsneeds("ontheblockneedscen") = True then %>
								 <table class="table table-responsive table-condensed">
									<tr>
										<td class="big"  style="text-align:center;width:33%;background-color: white"><span style="font-weight:bold;color:#9a1400;">CEN</span></td>
									</tr>
								</table>
								<% end if %>
								<% if objRSteamsneeds("ontheblockneedsfor") = True then %>
								 <table class="table table-responsive table-condensed">
									<tr>
										<td class="big" style="text-align:center;width:33%;background-color: white"><span style="font-weight:bold;color:#9a1400;">FOR</span></td>
									</tr>
								</table>
								<% end if %>
								<% if objRSteamsneeds("ontheblockneedsgua") = True then %>
								 <table class="table table-responsive table-condensed">
									<tr>
										<td class="big" style="text-align:center;width:33%;background-color: white"><span style="font-weight:bold;color:#9a1400;">GUA</span></td>
									</tr>
								</table>
								<% end if %>
								<% if objRSteamsneeds("ontheblockneedsfc") = True then %>
								 <table class="table table-responsive table-condensed">
									<tr>
										<td class="big" style="text-align:center;width:33%;background-color: white"><span style="font-weight:bold;color:#9a1400;">F-C</span></td>
									</tr>
								</table>
								<% end if %>
								<% if objRSteamsneeds("ontheblockneedsgf") = True then %>
								<table class="table table-responsive table-condensed">
									<tr>
										<td class="big" style="text-align:center;width:33%;background-color: white"><span style="font-weight:bold;color:#9a1400;">G-F</span></td>
									</tr>
								</table>
								<% end if %>
							</td>
							<%
								objRSteamsneeds.Close
							%>
							<% if IsNull (objRSteams("ontheblockcomments")) or objRSteams("ontheblockcomments") = ""  then%>
								<td class="big"></td>
							<%else%>
								<td class="big" style="width:34%"><span><textarea style="color:black;background-color:white;font-size:12px;" name="ontheblockscomments" disabled class="form-control" rows="10" id="ontheblockscomments"><%=objRSteams("ontheblockcomments")%></textarea></span></td>
							<%end if%>
						</tr>
						<%
							objRSteams.MoveNext
							count = count + 1
							Wend
						%>
						<%
							objRSteams.Close								
							objRSteams.Open "SELECT ontheblockall FROM tblowners WHERE ownerid = "& ownerid &" ", objConn
						%>
					</table>
					<table class="table table-custom-black  table-bordered table-condensed">
				<% if objRSteams.Fields("ontheblockall").Value = true then%>
					<tr>
						<th style="color:black;text-transform:capitalize;"  colspan="2" class="info">My Block <input type="checkbox" name="chkAllPlayers" checked onClick="toggle(this)" /> Check All</th>
					</tr>
				<% else %>
					<tr>
						<th style="color:black;text-transform:capitalize;"  colspan="2" class="info">My Block <input type="checkbox"  name="chkAllPlayers" onClick="toggle(this)" /> Check All</th>
					</tr>
					<%end if%>
					<%
						ownerid = session("ownerid")	
						if ownerid = "" then
							GetAnyParameter "var_ownerid", ownerid
						end if
						objRSontheBlock.Open "SELECT * FROM tblplayers WHERE ownerid = "& ownerid &" and rentalplayer = 0 order by lastName", objConn,3,3,1 
					%>
					<tr>
						<td class="big" colspan="2">
							<table class="table table-custom-black  table-bordered table-condensed" border="0" style="border-collapse" cellpadding="0" cellspacing="0">
								<%
								loopcnt = 0					
								While loopcnt < objRSontheBlock.RecordCount
								%>
								<% ontheblockind = objRSontheBlock.Fields("ontheBlock").Value %>
								<tr>
									<% if ontheblockind = True then %>
									<td class="big" style="width:50%;background-color:yellowgreen;color:black;text-align:left;"><input type="checkbox" checked name="chkPid" value="<%=objRSontheBlock.Fields("PID").Value & ";" & objRSontheBlock.Fields("lastName").Value & ", " & objRSontheBlock.Fields("firstName").Value & ";" &  objRSontheBlock.Fields("ontheBlock").Value%>">
										<%= left(objRSontheBlock.Fields("firstName").Value,1)%>.&nbsp;<%=objRSontheBlock.Fields("lastName").Value%></td>
									<% else %>
									<td class="big" style="width:50;text-align:left;%"><input type="checkbox" name="chkPid" value="<%=objRSontheBlock.Fields("PID").Value & ";" & objRSontheBlock.Fields("lastName").Value & ", " & objRSontheBlock.Fields("firstName").Value & ";" &  objRSontheBlock.Fields("ontheBlock").Value%>">
										<%= left(objRSontheBlock.Fields("firstName").Value,1)%>.&nbsp;<%=objRSontheBlock.Fields("lastName").Value%></td>
									<% end if %>
									<%
									 objRSontheBlock.MoveNext
									 loopcnt = loopcnt + 1
									%>
									<% ontheblockind = objRSontheBlock.Fields("ontheBlock").Value %>
									<% If loopcnt < objRSontheBlock.RecordCount Then %>
									<% if ontheblockind = True then %>
									<td class="big" style="width:50%;background-color:yellowgreen;color:black;text-align:left;"><input type="checkbox" checked name="chkPid" value="<%=objRSontheBlock.Fields("PID").Value & ";" & objRSontheBlock.Fields("lastName").Value & ", " & objRSontheBlock.Fields("firstName").Value & ";" &  objRSontheBlock.Fields("ontheBlock").Value%>">
										<%= left(objRSontheBlock.Fields("firstName").Value,1)%>.&nbsp;<%=objRSontheBlock.Fields("lastName").Value%></td>
									<% else %>
									<td class="big" style="width:50%;text-align:left;"><input type="checkbox" name="chkPid" value="<%=objRSontheBlock.Fields("PID").Value & ";" & objRSontheBlock.Fields("lastName").Value & ", " & objRSontheBlock.Fields("firstName").Value & ";" &  objRSontheBlock.Fields("ontheBlock").Value%>">
										<%= left(objRSontheBlock.Fields("firstName").Value,1)%>.&nbsp;<%=objRSontheBlock.Fields("lastName").Value%></td>
									<% end if %>
									<% else %>
									<td></td>
									<% end if %>
								</tr>
								<%
								 objRSontheBlock.MoveNext
								 loopcnt = loopcnt + 1
								 Wend
								 objRSontheBlock.Close
								%>
							</table>
						</td>
					</tr>
					<%
						objRSteamsneeds.Open "SELECT ontheblockneedscen, ontheblockneedsfor, ontheblockneedsgua, " &_
																				 "ontheblockneedsfc, ontheblockneedsgf, ontheblockcomments " &_
																	"FROM tblowners WHERE ownerid = " & ownerid & " ", objConn
					%>
					<tr>
						<th style="color:black;"  colspan="4" class="info">My Needs</th>
					</tr>
					<tr>
						<td class="big" colspan="2" style="width:5%;text-align:left">
						<% if objRSteamsneeds("ontheblockneedscen") = True then %>
							<input type="checkbox" checked name="chkCenters" value="Centers">&nbsp;Centers
						</td>
						<% else %>
							<input type="checkbox" name="chkCenters" value="Centers">&nbsp;Centers
						</td>
						<% end if %>
					</tr>
					<tr>
						<td class="big" colspan="2" style="width:5%;text-align:left">
						<% if objRSteamsneeds("ontheblockneedsfor") = True then %>
							<input type="checkbox" checked name="chkForwards" value="Forwards">&nbsp;Forwards
						</td>
						<% else %>
							<input type="checkbox" name="chkForwards" value="Forwards">&nbsp;Forwards
						</td>
						<% end if %>
					</tr>
					<tr>
						<td class="big" colspan="2" style="width:5%;text-align:left">
						<% if objRSteamsneeds("ontheblockneedsgua") = True then %>
							<input type="checkbox" checked name="chkGuards" value="Guards">&nbsp;Guards
						</td>
						<% else %>
							<input type="checkbox" name="chkGuards" value="Guards">&nbsp;Guards
						</td>
						<% end if %>
					</tr>
					<tr>
						<td class="big" colspan="2" style="width:5%;text-align:left">
						<% if objRSteamsneeds("ontheblockneedsfc") = True then %>
							<input type="checkbox" checked name="chkFC" value="Forward-Centers">&nbsp;Forward-Centers
						</td>
						<% else %>
							<input type="checkbox" name="chkFC" value="Forward-Centers">&nbsp;Forward-Centers
						</td>
						<% end if %>
					</tr>
					<tr>
						<td class="big" colspan="2" style="width:5%;text-align:left">
						<% if objRSteamsneeds("ontheblockneedsgf") = True then %>
							<input type="checkbox" checked name="chkGF" value="Guard-Forwards">&nbsp;Guard-Forwards
						</td>
						<% else %>
							<input type="checkbox" name="chkGF" value="Guard-Forwards">&nbsp;Guard-Forwards
						</td>
						<% end if %>
					</tr>
					<%  if myTradeInd = true then %> 
					<tr>
						<td class="big" colspan="2"><textarea maxlength="150" name="ontheblockscomments" class="form-control" rows="4" placeholder="Enter OTB Comments" id="ontheblockscomments"><%=objRSteamsneeds("ontheblockcomments")%></textarea>
							<br>
							<button type="submit" value="Update" name="Action" class="btn btn-trades btn-block"><i class="far fa-arrow-alt-circle-down"></i>&nbsp;Update My Block</button>
							<input type="checkbox" name="email_league" value="Yes">
							<span class="glyphicon glyphicon-envelope"></span>&nbsp;Email the League with this update!</td>
					</tr>
					<%end if%>
				</table>
			</div>
		</div>
	</div>
</form>