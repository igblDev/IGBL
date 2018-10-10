<%
	Set objConn	= Server.CreateObject("ADODB.Connection")
	objConn.Open Application("lineupstest_ConnectionString")

	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
								"Persist Security Info=False"
	Set objRSteams = Server.CreateObject("ADODB.RecordSet")
	Set objRSParms = Server.CreateObject("ADODB.RecordSet") 
	
	objRSteams.Open "SELECT * FROM tblOwners WHERE (((OwnerID)=" & ownerid & "))", objConn
	
	objRSParms.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
	wTradeDeadLine = objRSWork.Fields("param_date").Value
	objRSParms.Close
%>
</br>
<form action="dashboard.asp" name="editteam" id="Registration" method="POST">
<input type="hidden" name="action" value="Save Form Data">
<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-default">
				<table class="table table-striped table-custom-black table-responsive table-bordered table-condensed">
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-user"></span> </span>
									<input class="form-control required" name="txtownerName" value="<%=objRSteams.Fields("OwnerName")%>" size="40" placeholder="Enter Owner(s) Name" maxlength="35" type="text" tabindex="1">
								</div>
								<span class="help-block text-danger text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter First/Last Name</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-user"></span> </span>
									<input class="form-control required"  name="txtTeamName" value="<%=objRSteams.Fields("TeamName")%>" size="26" placeholder="Enter Team Name" maxlength="25" type="text" tabindex="1">
								</div>
								<span class="help-block text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter Team Name (25 - Max)</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"> <span class="input-group-addon"> <span class="glyphicon glyphicon-user"></span> </span>
									<input class="form-control required" name="txtShortName" value="<%=objRSteams.Fields("ShortName")%>" size="5" maxlength="5" placeholder="Enter Team Abbreviation" type="text" tabindex="1">
								</div>
								<span class="help-block text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter Short Team Name (5 - Max)</span> 
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"><span class="glyphicon glyphicon-lock"></span> </span>
									<input class="form-control required" name="txtUserName" value="<%=objRSteams.Fields("userid")%>" size="20" maxlength="20" placeholder="Enter User ID" type="text" tabindex="1">
								</div>
								<span class="help-block text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter IGBL UserID (20 - Max)</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-lock"></span> </span>
									<input class="form-control" name="txtPassword" value="<%=objRSteams.Fields("password")%>" size="20" placeholder="Enter Password" maxlength="20" type="password" tabindex="1">
								</div>
								<span class="help-block text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter IGBL Password</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"> <span class="input-group-addon"> <span class="glyphicon glyphicon-lock"></span> </span>
									<input class="form-control" name="confirmation" value="<%=objRSteams.Fields("passconf")%>" size="20" placeholder="Confirm Password" maxlength="20" type="password" tabindex="1">
								</div>
								<span class="help-block text-left"><span class="redText"><i class="fa fa-asterisk" aria-hidden="true"></i></span> Enter IGBL Password Conf</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"> <span class="input-group-addon"> <span class="glyphicon glyphicon-phone"></span> </span>
									<input class="form-control" name="txtCPhone" value="<%=objRSteams.Fields("cellphone")%>" size="20" placeholder="Enter Primary Contact Number" type="tel" tabindex="1">
								</div>
								<span class="help-block text-left">Enter Cell # Here</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-envelope"></span> </span>
									<input class="form-control" name="txtEmailH" value="<%=objRSteams.Fields("homeemail")%>" size="100" maxlength="100" placeholder="Enter Primary Email Address" type="text" tabindex="1">
								</div>
								<span class="help-block text-left">Enter Primary Email Here</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-envelope"></span> </span>
									<input class="form-control"  name="textMessages" value="<%=objRSteams.Fields("textMessages")%>" size="100" maxlength="100" placeholder="Enter SMS Address" type="text" tabindex="1">
								</div>
								<span class="help-block text-left">Enter SMS Address</span>
							</div>
						</td>
					</tr>
					<tr>
						<td>
							<div class="form-group">
								<div class="col-md-10 input-group input-group-md"><span class="input-group-addon"> <span class="glyphicon glyphicon-envelope"></span> </span>
									<input class="form-control" name="txtteamlogo" value="<%=objRSteams.Fields("TeamLogo")%>" size="80" placeholder="Enter Team Logo Url" type="text" tabindex="1">
								</div>
								<span class="help-block text-danger text-left">Insert Logo Url. <redIcon>Best Displayed When Square!</redIcon></span> 
							</div>
						</td>
					</tr>
					<tr>
						<td style="background-color: yellowgreen;">
							<div class="checkbox">
								<% if (wTradeDeadLine + 1 + 1/24) < now() then %> 
								<% else %>
									<%if objRSteams.Fields("acceptTradeOffers").Value = true then%>
										<label>
										<span style="color:#737373;vertical-align: sub;font-weight:bold"><input type="checkbox" name="acceptTradeOffers" checked value="<%=objRSteams.Fields("acceptTradeOffers")%>">&nbsp;RECEIVE TRADE OFFERS</span>
										</label>
									<%else%>
										<label>
										<span style="color:#737373;vertical-align: sub;font-weight:bold"><input type="checkbox" name="acceptTradeOffers" value="<%=objRSteams.Fields("acceptTradeOffers")%>">&nbsp;RECEIVE TRADE OFFERS</span>
										</label>
									<%end if%>									
								<%end if%>		
							</div>
						</td>
					</tr>
						<tr>
						<th class="text-center">SET ALERTS</th>
					</tr>
					<tr style="background-color:black;">
						<td style="font-weight:bold;text-align:left;">
							<div class="checkbox checkbox-sm">
								<%if objRSteams.Fields("receiveEmails").Value = true then%>
								<label>
									<span style="color:#9a1400;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveEmails" checked value="<%=objRSteams.Fields("receiveEmails")%>">&nbsp;RECEIVE ALERTS VIA EMAIL</span> 
								</label>
								<%else%>
								<label>
									<span style="color:#9a1400;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveEmails" value="<%=objRSteams.Fields("receiveEmails")%>">&nbsp;RECEIVE ALERTS VIA EMAIL</span> 
								</label>
								<%end if%>		
							</div>
							<div class="checkbox checkbox-sm">
								<%if objRSteams.Fields("receiveTexts").Value = true then%>
								<label>
									<span style="color:#9a1400;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveTexts" checked value="<%=objRSteams.Fields("receiveTexts")%>">&nbsp;RECEIVE ALERTS VIA TEXT</span> 
								</label>
								<%else%>
								<label>
									<span style="color:#9a1400;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveTexts" value="<%=objRSteams.Fields("receiveTexts")%>">&nbsp;RECEIVE ALERTS VIA TEXT</span> 
								</label>
								<%end if%>		
							</div>
						 <div class="checkbox checkbox-sm">
								<%if objRSteams.Fields("receiveFreeAgentAlerts").Value = true then%>
								<label>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveFreeAgentAlerts" checked value="<%=objRSteams.Fields("receiveFreeAgentAlerts")%>">&nbsp;FREE AGENT SIGNINGS</span> 
								</label>
								<%else%>
								<label>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveFreeAgentAlerts" value="<%=objRSteams.Fields("receiveFreeAgentAlerts")%>">&nbsp;FREE AGENT SIGNINGS</span> 
								</label>
								<%end if%>		
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveRentalAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveRentalAlerts" checked value="<%=objRSteams.Fields("receiveRentalAlerts")%>">&nbsp;PLAYER RENTALS</span> 
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveRentalAlerts" value="<%=objRSteams.Fields("receiveRentalAlerts")%>">&nbsp;PLAYER RENTALS</span> 
								<%end if%>
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveTradeAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveTradeAlerts" checked value="<%=objRSteams.Fields("receiveTradeAlerts")%>">&nbsp;TRADES<small><whiteText> (ACCEPTS, OFFERS, REJECTS & WITHDRAWALS)</whiteText></small></span>
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveTradeAlerts" value="<%=objRSteams.Fields("receiveTradeAlerts")%>">&nbsp;TRADES<small><whiteText> (ACCEPTS, OFFERS, REJECTS & WITHDRAWALS)</whiteText></small></span>
								<%end if%>	
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveWaiverAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveWaiverAlerts" checked value="<%=objRSteams.Fields("receiveWaiverAlerts")%>">&nbsp;WAIVERS RUN</span> 
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveWaiverAlerts" value="<%=objRSteams.Fields("receiveWaiverAlerts")%>">&nbsp;WAIVERS RUN</span> 
								<%end if%>
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveStaggerAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveStaggerAlerts" checked value="<%=objRSteams.Fields("receiveStaggerAlerts")%>">&nbsp;STAGGER WINDOW (O/C)</span> 
								<%else%>
								<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveStaggerAlerts" value="<%=objRSteams.Fields("receiveStaggerAlerts")%>">&nbsp;STAGGER WINDOW (O/C)</span> 
								<%end if%>
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveBoxScoreAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveBoxScoreAlerts" checked value="<%=objRSteams.Fields("receiveBoxScoreAlerts")%>">&nbsp;BOX SCORES GENERATED</span> 
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveBoxScoreAlerts" value="<%=objRSteams.Fields("receiveBoxScoreAlerts")%>">&nbsp;BOX SCORES GENERATED</span> 
								<%end if%>
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveOTBAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveOTBAlerts" checked value="<%=objRSteams.Fields("receiveOTBAlerts")%>">&nbsp;ON THE BLOCK POSTS</span> 
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveOTBAlerts" value="<%=objRSteams.Fields("receiveOTBAlerts")%>">&nbsp;ON THE BLOCK POSTS</span> 
								<%end if%>
								</label>
							</div>
							<div class="checkbox">
								<label>
								<%if objRSteams.Fields("receiveEmailLeagueAlerts").Value = true then%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveEmailLeagueAlerts" checked value="<%=objRSteams.Fields("receiveEmailLeagueAlerts")%>">&nbsp;EMAIL THE LEAGUE POSTS</span> 
								<%else%>
									<span style="color:yellowgreen;vertical-align: sub;font-weight:bold"><input type="checkbox" name="receiveEmailLeagueAlerts" value="<%=objRSteams.Fields("receiveEmailLeagueAlerts")%>">&nbsp;EMAIL THE LEAGUE POSTS</span> 
								<%end if%>
								</label>
							</div>
						</td>
					</tr>
				</table>
			</div>
		</div>
	</div>
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12" style="text-align:right;">
      <button type="submit" value="Save Record" class="btn btn-default"><i class="far fa-arrow-alt-circle-down"></i>&nbsp;Save Record</button>
      <button type="reset" value="Reset" name="Reset" class="btn btn-default"><i class="fas fa-sync-alt"></i>&nbsp;Refresh</button>
		</div>
	</div>
	<br>
</form>
