					<br>	
					<div class="row">					
						<div class="col-xs-4">
								<img class="img-responsive" style="max-width:100%;height:auto;margin:0px auto;display:block;border: #111;border-width: thin;border-style: solid;" src="<%=Logo%>">
						</div>
						<div class="col-xs-8">
							<table class="table table-custom-black table-condensed table-bordered table-responsive">
								<tr style="background-color:white;">									
									<td class="big" style="width: 50%;background-color: black;color: yellowgreen;font-weight: bold;">RK: <%=w_tm_rank%>&nbsp;|&nbsp;(<%=objRSteamsSession("won")%>-<%=objRSteamsSession("loss")%>)</td>
									<td class="big" style="width: 50%;background-color: yellowgreen;color: black;font-weight: bold;">Cycle <%=w_current_cycle%>&nbsp;-&nbsp;(<%=cycleWins%>-<%=cycleLoss %>)</td>
								</tr>
								<tr>									
									<td class="big" style="width: 50%;font-weight: bold;color:#01579B;background-color:#fcf8e3;">Waiver BAL: <span style="color:red;font-weight:bold;">$<%= w_WaiverBal%></span></td>
									<td class="txn" style="background-color:white"><a class="blue" href="transelect.asp">Acquisitions</a></td>
								</tr>
								<tr style="background-color:white">
									<td  width="50%" class="txn"><a class="blue" href="teamanalysis.asp">Team Analysis</a></td><td class="txn"><a class="blue" href="viewLineups.asp">Matchups</a></td>
								</tr>
								<%if gameStartedCnt > 0 then %>
								<tr style="background-color:white">
									<td colspan="2">
										<table class="table table-responsive table-condensed table-bordered" style="background-color:#ddd;">
											<tr>
												<td style="width:33%;text-align:center"><greenIcon2><i class="fas fa-basketball-ball  fa-spin"></i> Active<span class="sr-only"></span></greenIcon2></td>
												<td style="width:34%;text-align:center"><redIcon2><i class="fas fa-ban"></i> Banned</redIcon2></td>
												<td style="width:33%;text-align:center"><blackIcon2><i class="fas fa-user-lock"></i> Locked</blackIcon2></td>					
											</tr>
										</table>	
									</td>
								</tr>
								<% else %>
								<tr style="background-color:white">
									<% if (date() > wPlayoffStart) then %>
										 <td  style="width:50%" class="txn"><a class="blue" href="results.asp">Results</a></td><td class="txn"><a class="blue" href="allstandings.asp">Standings</a></td>
									<% else %>
										 <td  style="width:50%" class="txn"><a class="blue" href="results.asp">Results</a></td><td class="txn"><a class="blue" href="allstandings.asp">Standings</a></td>
									<% end if %>
								</tr>
								<tr style="background-color:white">
									<td class="txn"><a class="blue" href="playerSearch.asp">Player Search</a></td><td style="width:50%" class="txn"><a class="blue" href="allskeds.asp">Schedule</a></>
									</tr>
								<%end if %>
							</table>	
						</div>
					</div>
					<!--BUTTON LINKS-->
					</br>
					<div class="row">
						<div class="col-xs-12">
							<div class="btn-group btn-group-justified" >
							<div class="btn-group">
								<button class="btn btn-links big"  name="Action"><i class="far fa-file-invoice-dollar"></i>&nbsp;<a href="finreports.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;">$<%= w_total_spent %></a></button>
							</div>
							<div class="btn-group">					
								<% if  objRStradesreceived.RecordCount > 0 then %>
								<button class="btn btn-links big"  name="Action"><i class="far fa-inbox-in"></i>&nbsp;<a href="tradeoffers.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;"><%= objRStradesreceived.RecordCount %></a></button>
								<%else%>
								<button class="btn btn-links big"  name="Action"><span><i class="far fa-inbox-in"></i></span>&nbsp;0</button>
								<%end if%>
							</div>
							<div class="btn-group">				
								<% if  objRStradesmade.RecordCount > 0 then %>
								<button class="btn btn-links big"  name="Action"><i class="far fa-inbox-out"></i>&nbsp;<a href="pendingtrades.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;"><%= objRStradesmade.RecordCount %></a></button>
								<%else%>
								<button class="btn btn-links big"  name="Action"><span><i class="far fa-inbox-out"></i></span>&nbsp;0</button>
								<%end if%>
							</div>
							<div class="btn-group">
								<% if  objRSwaivers.RecordCount > 0 then %>
								<button class="btn btn-links big"  name="Action"><i class="far fa-cut"></i>&nbsp;<a href="pendingwaivers.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;"><%= objRSwaivers.RecordCount %></a></button>
								<%else%>
								<button class="btn btn-links big"  name="Action"><span><i class="far fa-cut"></i>&nbsp;0</button>
								<%end if%>
							</div>
							<div class="btn-group">
								<% if  objrsForecast.RecordCount > 0 then %>
								<button class="btn btn-links big"  name="Action"><i class="fa fa-balance-scale"></i>&nbsp;<a href="pendingAnalyzedTrades.asp?ownerid=<%= ownerid %>" style="color: white;text-decoration: underline;"><%= objrsForecast.RecordCount %></a></button>
								<%else%>
								<button class="btn btn-links big"  name="Action"><span><i class="fa fa-balance-scale"></i></span>&nbsp;0</button>
								<%end if%>
							</div>
							</div>
						</div>
					</div>
					<form action="dashboard.asp" method="POST" onSubmit="return functionLineup(this)" name="frmLineups"  language="JavaScript">
					<input type="hidden" name="var_ownerid"  value="<%= ownerid %>" />
					<input type="hidden" name="var_gameTime" value="<%= gameDeadline %>" />
					<input type="hidden" name="var_gameTimeStagger" value="<%= gameStaggerDeadline %>" />
					<input type="hidden" name="selectDate" value="<%= selectDate %>" />
					<input type="hidden" name="centerCnt" value="<%= centerCnt %>" />
					<input type="hidden" name="forwardCnt" value="<%= forwardCnt %>" />
					<input type="hidden" name="guardCnt" value="<%= guardCnt %>" />	
					<input type="hidden" name="cDeadlinePassed"  value="<%= cDeadlinePassed %>" />
					<input type="hidden" name="f1DeadlinePassed" value="<%= f1DeadlinePassed %>" />
					<input type="hidden" name="f2DeadlinePassed" value="<%= f2DeadlinePassed %>" />	
					<input type="hidden" name="g1DeadlinePassed" value="<%= g1DeadlinePassed %>" />
					<input type="hidden" name="g2DeadlinePassed" value="<%= g2DeadlinePassed %>" />
					<input type="hidden" name="selectedCenter" value="<%= selectedCenter %>" />
					<input type="hidden" name="selectedForward1" value="<%= selectedForward1 %>" />
					<input type="hidden" name="selectedForward2" value="<%= selectedForward2 %>" />
					<input type="hidden" name="selectedGuard1" value="<%= selectedGuard1 %>" />
					<input type="hidden" name="selectedGuard2" value="<%= selectedGuard2 %>" />
					</br>
					<div class="row">
						<div class="col-xs-6">
						<select class="form-control " name="gameDays">
							<%if sAction = "Retrieve Lineup" then%>
							<option value="<%=selectDate%>"- selected><%=selectDate%></option>
							<% else %>
							<option value="<%=objRSgames.Fields("gameDay")%>" selected><%=objRSgames.Fields("gameDay")%></option>
							<% end if %>
							<% While not objRSgames.EOF %>
							<option value="<%=objRSgames("gameDay")%>"><%=objRSgames.Fields("gameDay")%> </option>
							<% objRSgames.MoveNext
								Wend 
							%>
						</select>
					</div>
					<div class="col-xs-6">
						<input class="form-control hidden" name="clickTest" value="" type="radio" tabindex="1">
						<input class="form-control hidden" name="clickTest" value="" type="radio" tabindex="1">
						<button class="btn btn-default  btn-block" type="submit"  id="RetrieveLineup" onclick="document.FrontPage_Form1.clickTest[0].checked = true;" value="Retrieve Lineup" name="Action"><span class="glyphicon glyphicon-search"></span>&nbsp;Retrieve Lineup</button>
					</div>
				</div>
				</br>
				<%  							   
				objRSWork.Open "SELECT * FROM tblplayers " & _
                               "where ownerid = "&ownerId&" AND " & _
                               "(ir <> 0 or pendingTrade <> 0 or pendingwaiver <> 0 or rentalplayer <> 0) ", objConn,3,3,1
							   							   
                if objRSWork.RecordCount > 0 then							   
			   %>
				  <div class="row">
					  <div class="col-xs-12">
						  <div style="text-align: center;"><strong><mark>Legend:</mark></strong>&nbsp;&nbsp;<i class="fas fa-briefcase-medical red"></i>&nbsp;<strong>IR</strong>&nbsp;&nbsp;<i class="far fa-exchange"></i>&nbsp;<strong>Pending Trade</strong>&nbsp;&nbsp;<i class="fas fa-user-clock auctionText"></i>&nbsp;<strong>Pending Waiver</strong>&nbsp;&nbsp;<i class="fa fa-registered blue" aria-hidden="true"></i>&nbsp;<strong>Rented</strong></div>
					  </div>
				  </div>
				  </br>
			 <% end if 
			    objRSWork.Close %>
				<!--TRADELINE DEADLINE BANNER-->
				<% if (date() >= (wTradeDeadLine - 10)) AND (wTradeDeadLine >= date()) then %>
					<div class="row">		
						<div class="col-xs-12">
							<div class="alert alert-danger">
								<strong><i class="fa fa-exclamation-triangle" aria-hidden="true"></i>Trade-Deadline: </strong> <%=wTradeDeadLine%>
							</div>
						</div>
					</div>
				<% end if %>  				
				<% 
				objRSLU.Close	
				if centerCnt > 0 Then 
				%>	
					<div class="row">
						<div class="col-xs-12">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<th class="text-uppercase text-left;" colspan="2">CENTER</th>
									<th style="text-align:center;">AVG</th>
									<th style="text-align:center;">L/5</th>
									<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
								</tr>
								<%
									While Not objRScen.EOF
									 if len(objRScen.Fields("GameTime").Value) = 10 then
											wtime = Left(objRScen.Fields("GameTime").Value,4) & Right(objRScen.Fields("GameTime").Value,3)
									 else
											wtime = Left(objRScen.Fields("GameTime").Value,5) & Right(objRScen.Fields("GameTime").Value,3)
									 end if
								%>
								<% if cDeadlinePassed = true then%>
									<% if objRScen.Fields("PID").Value = startingCenterPID then %>
										<tr class="success">
											<!--#include virtual="Common/centerLineup.inc"-->			
											<td style="vertical-align:middle;text-align:center;width:10%"><greenIcon> <i class="fas fa-basketball-ball  fa-spin"></i></greenIcon><span class="sr-only"></span></td>
										</tr>
									<% else %>
										<tr style="background-color:white;">	
											<!--#include virtual="Common/centerLineup.inc"-->			
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon><i class="fas fa-ban"></i></redIcon></td>
										</tr>
									<% end if%>
								<% elseif  ((CDATE(objRScen.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then %>
									<tr style="background-color:white;">	
										<!--#include virtual="Common/centerLineup.inc"-->			
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>									
									</tr>
								<% elseif objRScen.Fields("PID").Value = startingCenterPID then%>							
									<tr style="background-color:white;">	
										<!--#include virtual="Common/centerLineup.inc"-->			
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sCenter" value="<%=objRScen.Fields("firstName").Value & " " & objRScen.Fields("lastName").Value & ";" & objRScen.Fields("PID").Value & ";" & objRScen.Fields("barps").Value  & ";" &  objRScen.Fields("GameTime").Value & ";"%>"></td>
									</tr>
								<% else %>
									<tr style="background-color:white;">	
										<!--#include virtual="Common/centerLineup.inc"-->							
										<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sCenter" value="<%=objRScen.Fields("firstName").Value & " " & objRScen.Fields("lastName").Value & ";" & objRScen.Fields("PID").Value & ";" & objRScen.Fields("barps").Value  & ";" &  objRScen.Fields("GameTime").Value & ";"%>"></td>
									</tr>
								<% end if %>
								<%
									objRScen.MoveNext
									Wend
								%>
							</table>
							<br>
						</div>
					</div>				
				<%else%>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown"></i>&nbsp;No Centers Available Tonight!
							</div>
						</div>
					</div>
				<% end if 
					 objRScen.close
				%>				
				<% if forwardCnt > 0 Then %>
					<div class="row">
						<div class="col-xs-12">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<th class="text-uppercase text-left;" colspan="2">FORWARDS</th>
									<th style="text-align:center;">AVG</th>
									<th style="text-align:center;">L/5</th>
									<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
								</tr>
									<%
									While Not objRSfor.EOF
										 if len(objRSfor.Fields("GameTime").Value) = 10 then
												wtime = Left(objRSfor.Fields("GameTime").Value,4) & Right(objRSfor.Fields("GameTime").Value,3)
										 else
												wtime = Left(objRSfor.Fields("GameTime").Value,5) & Right(objRSfor.Fields("GameTime").Value,3)
										 end if							   
										 objRSNext5.Open  	"SELECT * FROM qryAllPlayerGameDays WHERE pid = " & objRSfor.Fields("PID").Value & " and gameday >= Date() order by gameday ", objConn,1,1	
										%>
									<% if f1DeadlinePassed AND f2DeadlinePassed then%>
										<% if ((objRSfor.Fields("PID").Value = selectedForward1) or (objRSfor.Fields("PID").Value = selectedForward2)) then %>
											<tr class="success">
												<!--#include virtual="Common/forwardLineup.inc"-->
												<td style="vertical-align:middle;text-align:center;width:10%"><greenIcon> <i class="fas fa-basketball-ball  fa-spin"></i></greenIcon><span class="sr-only"></span></td>
											</tr>	
										<% else %>
											<tr style="background-color:white;">	
													<!--#include virtual="Common/forwardLineup.inc"-->
												<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon><i class="fas fa-ban"></i></redIcon></td>
											</tr>	
										<% end if%>		
									<% elseif ((objRSfor.Fields("PID").Value = selectedForward1) or (objRSfor.Fields("PID").Value = selectedForward2))then %>
										<tr class="success">
										<!--#include virtual="Common/forwardLineup.inc"-->
											<td style="vertical-align:middle;text-align:center;width:10%"><greenIcon> <i class="fas fa-basketball-ball  fa-spin"></i></greenIcon><span class="sr-only"></span></td>
										</tr>						
									<% elseif ((CDATE(objRSfor.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then %>
										<tr style="background-color:white;">
										<!--#include virtual="Common/forwardLineup.inc"-->
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>									
										</tr>
									<% elseif ((objRSfor.Fields("PID").Value = startingForwardPID) or (objRSfor.Fields("PID").Value = startingForward2PID)) then%>
										<tr style="background-color:white;">
										<!--#include virtual="Common/forwardLineup.inc"-->
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sForward" value="<%=objRSfor.Fields("firstName").Value & " " & objRSfor.Fields("lastName").Value & ";" & objRSfor.Fields("PID").Value & ";" & objRSfor.Fields("barps").Value  & ";" &  objRSfor.Fields("GameTime").Value & ";"%>"></td>
										</tr>
									<% else %>
										<tr style="background-color:white;">
										<!--#include virtual="Common/forwardLineup.inc"-->
											<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sForward" value="<%=objRSfor.Fields("firstName").Value & " " & objRSfor.Fields("lastName").Value & ";" & objRSfor.Fields("PID").Value & ";" & objRSfor.Fields("barps").Value  & ";" &  objRSfor.Fields("GameTime").Value & ";"%>"></td>
										</tr>
									<% end if %>
									<%
										objRSfor.MoveNext
										Wend
									%>
							</table>
							<br>
						</div>
					</div>
					<% else %>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown"></i>&nbsp;No Forwards Available Tonight!
							</div>
						</div>
					</div>
					<% end if 
						 objRSfor.close
					%>
				<!--GUARDS-->					
				<% if guardCnt > 0 Then %>
				<div class="row">
					<div class="col-xs-12">
						<table class="table table-custom-black table-bordered table-responsive table-condensed">
							<tr>
								<th class="text-uppercase text-left;" colspan="2">GUARDS</th>
								<th style="text-align:center;">AVG</th>
								<th style="text-align:center;">L/5</th>
								<th style="text-align:center;"><i class="fas fa-basketball-hoop"></i></th>
							</tr>
						<%
						'Response.Write "guardCnt = " & guardCnt  & ".<br>"
						While Not objRSgua.EOF
						 'Response.Write "Inside loop for " & objRSgua.Fields("lastName").Value  & ".<br>"
						 if len(objRSgua.Fields("GameTime").Value) = 10 then
								wtime = Left(objRSgua.Fields("GameTime").Value,4) & Right(objRSgua.Fields("GameTime").Value,3)
						 else
								wtime = Left(objRSgua.Fields("GameTime").Value,5) & Right(objRSgua.Fields("GameTime").Value,3)
						 end if
						%>
						<% if g1DeadlinePassed  AND g2DeadlinePassed then%>
							<% if ((objRSgua.Fields("PID").Value = selectedGuard1) or (objRSgua.Fields("PID").Value = selectedGuard2)) then %>
						<tr class="success">
							<!--#include virtual="Common/guardLineup.inc"-->
								<td style="vertical-align:middle;text-align:center;width:10%"><greenIcon> <i class="fas fa-basketball-ball  fa-spin"></i></greenIcon><span class="sr-only"></span></td>							
						</tr>	
							<% else %>
						<tr style="background-color:white;">						
							<!--#include virtual="Common/guardLineup.inc"-->
							<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><redIcon><i class="fas fa-ban"></i></redIcon></td>
						</tr>
							<% end if%>
					<% elseif ((objRSgua.Fields("PID").Value = selectedGuard1) or (objRSgua.Fields("PID").Value = selectedGuard2))then %>
						<tr class="success">						
							<!--#include virtual="Common/guardLineup.inc"-->
							<td style="vertical-align:middle;text-align:center;width:10%"><greenIcon><i class="fas fa-basketball-ball  fa-spin"></i></greenIcon><span class="sr-only"></span></td>
						</tr>	
					<% elseif ((CDATE(objRSgua.Fields("GameTime").Value) <  CDATE(lineupTimeChk)) and (CDATE(selectDate) = date()))then %>
						<tr style="background-color:white;">						
							<!--#include virtual="Common/guardLineup.inc"-->
							<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><blackIcon><i class="fas fa-user-lock"></i></blackIcon></td>
						</tr>
					<% elseif ((objRSgua.Fields("PID").Value = startingGuardPID)  or (objRSgua.Fields("PID").Value = startingGuard2PID)) then %>
						<tr style="background-color:white;">							
							<!--#include virtual="Common/guardLineup.inc"-->
							<td  style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" checked id="C" name="sGuard" value="<%=objRSgua.Fields("firstName").Value & " " & objRSgua.Fields("lastName").Value & ";" & objRSgua.Fields("PID").Value & ";" &  objRSgua.Fields("barps").Value  & ";" &  objRSgua.Fields("GameTime").Value & ";"%>"></td>
						</tr>
					<% else %>
						<tr style="background-color:white;">						
							<!--#include virtual="Common/guardLineup.inc"-->
							<td style="vertical-align:middle;text-align:center;background-color:white;width:10%"><input type="checkbox" id="C" name="sGuard" value="<%=objRSgua.Fields("firstName").Value & " " & objRSgua.Fields("lastName").Value & ";" & objRSgua.Fields("PID").Value & ";" & objRSgua.Fields("barps").Value  & ";" &  objRSgua.Fields("GameTime").Value & ";"%>"></td>
						</tr>
					<% end if %>
					<%
						objRSgua.MoveNext
						Wend
					%>
					</table>
				</div>
			</div>
			<% else %>
					<div class="row">
						<div class="col-xs-12">
							<div class="jumbotron" style="background-color:#9a1400;color:white;text-align:center;vertical-align:middle">
								<i class="fas fa-frown"></i>&nbsp;No Guards Available Tonight!
							</div>
						</div>
					</div>
			<% end if 
				 objRSgua.close
			%>	
			</br>
			<div class="row">
				<div class="col-xs-12">
					<%if cDeadlinePassed = false or f1DeadlinePassed = false or f2DeadlinePassed = false or g1DeadlinePassed = false or g2DeadlinePassed = false then%>
						<button type="submit" id="idSubmitLineup"  value="Submit Lineup" name="Action" class="btn   btn-block btn-default"><span class="glyphicon glyphicon-save"></span>&nbsp;Submit Lineup</button>
					<%end if%>
				</div>			
			</div>			
			</br>
			<% if unAvailCnt > 0 then %>
				<div class="row">
					<div class="col-xs-12">	
						<div class="panel panel-override">
							<table class="table table-custom-black table-bordered table-responsive table-condensed">
								<tr>
									<td style="border-radius: unset;vertical-align:middle;text-align:center;background-color:black;color:white;text-transform:uppercase;font-weight:bold;" colspan="4"><i class="fas fa-power-off"></i>&nbsp;<%= (FormatDateTime(selectDate,1)) %></td>
								</tr>
								<tr>
									<th style="width:60%;" class="text-uppercase text-center">Player</th>
									<th style="text-align:center;width:15%;">AVG</th>
									<th style="text-align:center;width:15%;">L/5</th>
									<th style="text-align:center;width:10%;">USG</th>
								</tr>
								<%
									While Not objRSUnavail.EOF
								%>
								<tr style="background-color:white">
									<td class="big" style="width:55%;"><a class="blue" href="playerprofile.asp?pid=<%=objRSUnavail.Fields("PID").Value %>"><%=left(objRSUnavail.Fields("firstName").Value,14)%>&nbsp;<%=left(objRSUnavail.Fields("lastName").Value,14)%></a>
									<br><small><span class="greenTrade text-uppercase"><%=objRSUnavail.Fields("TeamName").Value%></span>&nbsp;|&nbsp;<span class="orange"><%=objRSUnavail.Fields("pos").Value%></span></small>
									<% if objRSUnavail.Fields("IR").Value = true then %>
									<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
									<%end if%>
									<% if objRSUnavail.Fields("pendingtrade").Value = true then %>
									<strong>	<i class="far fa-exchange"></i></strong>
									<%end if%>
										<% if objRSUnavail.Fields("pendingWaiver").Value = true then %>
									<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
										<%end if%>									
									</td>					
									<td style="width:15%;text-align:center;vertical-align:middle;" class="text-uppercase big"><span class="badgeBlue"><%=round(objRSUnavail.Fields("barps").Value,2)%></span></td>		
										<% if CDbl(objRSUnavail.Fields("l5barps").Value) > CDbl(objRSUnavail.Fields("barps").Value) then %>
											<td style="vertical-align:middle;text-align:center"><span class="badgeUp big"><%= round(objRSUnavail.Fields ("l5barps").Value,2) %></span></td>
										<% elseif CDbl(objRSUnavail.Fields("barps").Value) > CDbl(objRSUnavail.Fields("l5barps").Value) then%>
											<td style="vertical-align:middle;text-align:center"><span class="badgeDown big"><%= round(objRSUnavail.Fields ("l5barps").Value,2) %></span></td>
										<%else%>
										<td class="big" style="vertical-align:middle;text-align:center"><span class="badgeEven"><%= round(objRSUnavail.Fields ("l5barps").Value,2) %></span></td>
										<%end if %>	

										<%if objRSUnavail.Fields ("usage").Value > 0 then %>
											<td class="big" style="vertical-align:middle;width:15%;text-align:center;"class="text-uppercase"><span class="badgeUsage big"><%= round(objRSUnavail.Fields ("usage").Value,0) %></span></td>
										<%else %>
											<td class="big" style="vertical-align:middle;width:15%;text-align:center;"class="text-uppercase"><span class="badgeUsage big">0</span></td>
										<%end if%>
									</td>					
								</tr>
								<%
								objRSUnavail.MoveNext					
								Wend
								%>
							</table>
						</div>
					</div>
				</div>
			<%end if%>
			</form>	
		<%end if
		if sAction = "Submit Lineup"  and errorCode = "" then %>
		<!--#include virtual="Common/headerMain.inc"-->
		<form action="dashboard.asp" method="POST" name="frmconfirm">
		<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
		<input type="hidden" name="txtOwnerID" value="<%= ownerid %>" />
		<%
		Dim objRSLineupsPics  

		Set objRSLineupsPics = Server.CreateObject("ADODB.RecordSet")
		
		objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wCenterPID  & " = PID " , objConn,3,3,1
		cFirstName = objRSLineupsPics.Fields("firstName").Value
		cLastName  = objRSLineupsPics.Fields("lastName").Value
		objRSLineupsPics.Close
		
		objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wForwardPID  & " = PID " , objConn,3,3,1
		f1FirstName= objRSLineupsPics.Fields("firstName").Value
		f1LastName = objRSLineupsPics.Fields("lastName").Value
		objRSLineupsPics.Close
		
		objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " &  wForward2PID  & " = PID " , objConn,3,3,1
		f2FirstName= objRSLineupsPics.Fields("firstName").Value
		f2LastName = objRSLineupsPics.Fields("lastName").Value
		objRSLineupsPics.Close
		
		objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wGuardPID  & " = PID " , objConn,3,3,1
		g1FirstName= objRSLineupsPics.Fields("firstName").Value
		g1LastName = objRSLineupsPics.Fields("lastName").Value
		objRSLineupsPics.Close
		
		objRSLineupsPics.Open "Select firstName, lastName from tblplayers where  " & wGuard2PID & " = PID " , objConn,3,3,1
		g2FirstName= objRSLineupsPics.Fields("firstName").Value
		g2LastName = objRSLineupsPics.Fields("lastName").Value
		objRSLineupsPics.Close 
	%>
	<% sAction = "Return"%>
		<div class="container">
			<div class="row">
				<div class="col-xs-6">
					<select class="form-control " name="gameDays">
						<%if sAction = "Retrieve Lineup" then%>
						<option value="<%=selectDate%>" selected><%=selectDate%>  - Next Game Date</option>
						<% else %>
						<option value="<%=objRSgames.Fields("gameDay")%>" selected><%=objRSgames.Fields("gameDay")%> - Next Game Date</option>
						<% end if %>
						<% While not objRSgames.EOF %>
						<option value="<%=objRSgames("gameDay")%>"><%=objRSgames.Fields("gameDay")%></option>
						<% objRSgames.MoveNext
						Wend 
						%>
					</select>
				</div>
				<div class="col-xs-6 align="right">
					<input class="form-control hidden" name="clickTest" value="" type="radio" tabindex="1">
					<input class="form-control hidden" name="clickTest" value="" type="radio" tabindex="1">
					<button type="submit"	onclick="document.FrontPage_Form1.clickTest[0].checked = true;" value="Retrieve Lineup" name="Action" class="btn btn-default  btn-block"><span class="glyphicon glyphicon-search"></span>&nbsp;Retrieve Lineup</button>
				</div>
			</div>
			<br>
			<div class="row">
				<div class="col-xs-12">
					 <div class="panel panel-override">
							<table class="table table-custom-black table-striped table-bordered table-condensed">
								<tr>
								<% if date() = cdate(gamedate) then %>
									<th colspan="2" class="panel-title">Today's&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;<%= gamedate %></th>
								<%else%>
									<th colspan="2" class="panel-title">Future&nbsp;<i class="fal fa-calendar-alt"></i>&nbsp;<%= gamedate %></th>
								<%end if%>
								<tr>
								<tr bgcolor="#FFFFFF">
									<td class="big" colspan="2" style="text-align:center;"><%=cFirstName%>&nbsp;<%=cLastName%>&nbsp;<orangePos>(C)</orangePos></td>
								</tr>
								<tr bgcolor="#FFFFFF">
									<td class="big" width="50%" style="text-align:left"><%=f1FirstName%>&nbsp;<%=f1LastName%>&nbsp;<orangePos>(F)</orangePos></td>
									<td class="big" width="50%" style="text-align:right"><%=f2FirstName%>&nbsp;<%=f2LastName%>&nbsp;<orangePos>(F)</orangePos></td>
								</tr>
								<tr bgcolor="#FFFFFF">
									<td class="big" width="50%" style="text-align:left"><%=g1FirstName%>&nbsp;<%=g1LastName%>&nbsp;<orangePos>(G)</orangePos></td>
									<td class="big" width="50%" style="text-align:right"><%=g2FirstName%>&nbsp;<%=g2LastName%>&nbsp;<orangePos>(G)</orangePos></td>
								</tr>
							</table>
					</div>
				</div>
			</div>
			<div class="row">
				<div class="col-xs-12">
					<span><a href="viewLineups.asp" class="btn btn-block  btn-default" style="min-height: 40px;min-width: 40px;">View All Lineups & Matchups</a></span>
				</div>
			</div>
		</div>
		</form>
		<%end if
		if sAction ="Submit Lineup" and tip_time_error_flag = 1  then %>
		<!--#include virtual="Common/headerMain.inc"-->
		<form action="dashboard.asp" method="POST" name="frmreject" language="JavaScript">
		<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
		<div class="container">
			<div class="row">
				<div class="col-xs-12">
					<div class="alert alert-danger">
					 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
						<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i></strong><br> <%=errorcode%> 
					</div>
					<div>
						<button type="submit" value="" name="" class="btn btn-block btn-default"><i class="fas fa-trash-alt"></i>&nbsp;Return to Line-ups!</button>
					</div>
				</div>
			</div>
		</div>
		</form>
		<% elseif sAction ="Submit Lineup" and errorcodeLU = "Invalid Line-up" then  %>
		<!--#include virtual="Common/headerMain.inc"-->
		<form action="dashboard.asp" method="POST" name="frmreject" language="JavaScript">
		<input type="hidden" name="var_ownerid" value="<%= ownerid %>" />
		<div class="container">
			<div class="row">
				<div class="col-xs-12">
					<div class="alert alert-danger">
					 <a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
						<strong><i class="fa fa-exclamation-triangle fa-lg" aria-hidden="true"></i>&nbsp;Line-Up Error!</strong><br><br> <%=errorcode%> 
					</div>
					<div>
					 <a href="#" onClick="history.go(-1); return false;"> <button class="btn btn-block btn-default"><i class="fas fa-trash-alt"></i>&nbsp;Return to Line-ups!</button></a>
					</div>
				</div>
			</div>
		</div>
		</form>
		<%
		end if
		objRSNext5.Close
		objRSflex.Close
		objRSUnavail.Close
		objRSteamRec.Close
		objRSAll.Close
		Set objRSteamRec = Nothing
		ObjConn.Close
		Set objConn = Nothing
		Session.CodePage = Session("FP_OldCodePage")
		Session.LCID = Session("FP_OldLCID")
		%>
<!--END of TAB 1-->