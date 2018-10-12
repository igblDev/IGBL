					<tr>
						<%if objrsPlayerPos.Fields("lastTeamInd").Value = ownerId then %>
								<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;"><i class="fas fa-user-lock"></i></td>
								<td>
								<a class="blue" href="playerprofile.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=objrsPlayerPos.Fields("firstName").Value%>&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
								<br><small><span class="greenTrade"><%= objrsPlayerPos.Fields("team").Value%></span>&nbsp;<span class="orange"><%=objrsPlayerPos.Fields("pos").Value%></span>
								</td>
						<%else%>
							<% if objrsPlayerPos.Fields("playerStatus").Value = "W" then %>
							<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;">
								<button type="submit" value="Waiver Claim;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"        name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw fa-lg"></i></button>
							</td>
								<%elseif objrsPlayerPos.Fields("playerStatus").Value = "F" and objRSToday.RecordCount > 0 then %>
									<%if objRSNBASked.RecordCount > 0 then %>
										<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;">
											<button type="submit" value="Sign Free Agent;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"   name="Action" class="btn justify-content-center align-items-center btn-txn-green btn-xs"><i class="fa fa-plus fa-fw fa-lg"></i></button>
											<button type="submit" value="Rent Player(s);<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"    name="Action" class="btn justify-content-center align-items-center btn-txn-blue btn-xs"><i class="fa fa-registered fa-lg" aria-hidden="true"></i></button>
										</td>
									<%else%>
										<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;">
											<button type="submit" value="Sign Free Agent;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"   name="Action" class="btn justify-content-center align-items-center btn-txn-green btn-xs"><i class="fa fa-plus fa-fw fa-lg"></i></button>
										</td>
									<%end if%>
								<%elseif objrsPlayerPos.Fields("playerStatus").Value = "S" then %>
									<%if objRSNBASked.RecordCount > 0 AND wTipTime > (time() - 1/24) then %>
									<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;">
										<button type="submit" value="Waiver Claim;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"      name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw "></i></button>
										<button type="submit" value="Rent Player(s);<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"    name="Action" class="btn justify-content-center align-items-center  btn-txn-blue btn-xs"><strong><i class="fa fa-registered fa-lg" aria-hidden="true"></i></strong></button>
									</td>
									<%else%>
									<td class ="big"  style="white-space:nowrap;vertical-align:middle;text-align:center;">
										<button type="submit" value="Waiver Claim;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"      name="Action" class="btn justify-content-center align-items-center btn-txn-red btn-xs"><i class="fa fa-plus fa-fw fa-lg"></i></button>
									</td>
									<%end if%>
							<%else%>
							<td class ="big" style="white-space:nowrap;vertical-align:middle;text-align:center;">
								<button type="submit" value="Sign Free Agent;<%=objrsPlayerPos.Fields("PID").Value & ";" & objrsPlayerPos.Fields("firstName").Value & " " & objrsPlayerPos.Fields("lastName").Value & ";" &  objrsPlayerPos.Fields("StatusDesc").Value%>"     name="Action" class="btn justify-content-center align-items-center btn-txn-green btn-xs"><i class="fa fa-plus fa-fw fa-lg"></i></button>
							</td>
						<% end if %>
						<% if objRSNBASked.RecordCount > 0 then %>
							<td class ="big" style="vertical-align:middle;text-align:left;"><!--<span class="auctionText">$<%=objrsPlayerPos.Fields("auctionPrice").Value %></span>&nbsp;-->
							<%if (len(objrsPlayerPos.Fields("firstName").Value) + len(objrsPlayerPos.Fields("lastName").Value)) >= 15 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=left(objrsPlayerPos.Fields("firstName").Value,1)%>.&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=objrsPlayerPos.Fields("firstName").Value%>&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%end if%>
							<br><small><span class="greenTrade"><%= objrsPlayerPos.Fields("team").Value%></span>&nbsp;<span class="orange"><%=objrsPlayerPos.Fields("pos").Value%></span>
							<br><span class="gameTip"><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<i class="far fa-clock"></i>&nbsp;<%=wtime%></small></span>
						<% else %>
							<td class ="big" style="vertical-align:middle;text-align:left;">
							<%if (len(objrsPlayerPos.Fields("firstName").Value) + len(objrsPlayerPos.Fields("lastName").Value)) >= 15 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=left(objrsPlayerPos.Fields("firstName").Value,1)%>.&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=objrsPlayerPos.Fields("firstName").Value%>&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%end if%>
							<br><small><span class="greenTrade"><%= objrsPlayerPos.Fields("team").Value%></span>&nbsp;<span class="orange"><%=objrsPlayerPos.Fields("pos").Value%></small></span>
						<%end if%>		
						</td>	
						<%end if%>
						<td class ="big" style="vertical-align:middle;text-align:center;" class="text-center">	
							<span class="badgeBlue big"><%= round(objrsPlayerPos.Fields ("barps").Value,2) %></span>
						</td>
				
						<% if CDbl(objrsPlayerPos.Fields("l5barps").Value) > CDbl(objrsPlayerPos.Fields("barps").Value) then %>
						<td style="vertical-align:middle;text-align:center;"><span class="badgeUp big"><%= round(objrsPlayerPos.Fields ("l5barps").Value,2) %></span></td>
						<% elseif CDbl(objrsPlayerPos.Fields("barps").Value) > CDbl(objrsPlayerPos.Fields("l5barps").Value) then%>
						<td style="vertical-align:middle;text-align:center;"><span class="badgeDown big"><%= round(objrsPlayerPos.Fields ("l5barps").Value,2) %></span></td>
						<%else%>
						<td style="vertical-align:middle;text-align:center;"><span class="badgeEven big"><%= round(objrsPlayerPos.Fields ("l5barps").Value,2) %></span></td>
						<%end if %>
						</td>							
					</tr>	