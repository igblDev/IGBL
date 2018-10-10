							<td class="big" style="vertical-align:middle;text-align:center;vertical-align:bottom;width:20%;padding: inherit;"><img class="img-fluid center-block" src="<%= (objRSgua.Fields("image").Value)%>"></td>
							<td class="big" style="vertical-align:middle;text-align:left;width:34%;">
							<%if (len(objRSgua.Fields("firstName").Value) + len(objRSgua.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSgua.Fields("PID").Value %>">
									<%=left(objRSgua.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSgua.Fields("lastName").Value,14)%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSgua.Fields("PID").Value %>">
									<%=objRSgua.Fields("firstName").Value%>&nbsp;<%=objRSgua.Fields("lastName").Value%></a>
							<%end if%>		
							<br><small><span class="greenTrade"><%= objRSgua.Fields("team")%></span> | <span class="orange"><%= objRSgua.Fields("pos")%></span>	
							<% if objRSgua.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							<%if objRSgua.Fields("IR").Value = true then %>
								<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
							<%end if%>	
							<% if objRSgua.Fields("pendingtrade").Value = true then %>
								<strong>	<i class="far fa-exchange"></i></strong>
							<%end if%>
							<% if objRSgua.Fields("pendingWaiver").Value = true then %>
								<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
							<%end if%>
							<br><span class="gameTip"><%= objRSgua.Fields("opponent")%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%= wtime%></small>
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;width:18%;" class="text-center">
								<span class="badgeBlue"><%= round(objRSgua.Fields ("barps").Value,2) %></span>
							</td>
							<% if CDbl(objRSgua.Fields("l5barps").Value) > CDbl(objRSgua.Fields("barps").Value) then %>
								<td class="big" style="vertical-align:middle;text-align:center"><span class="badgeUp"><%= round(objRSgua.Fields ("l5barps").Value,2) %></span>
							<% elseif CDbl(objRSgua.Fields("barps").Value) > CDbl(objRSgua.Fields("l5barps").Value) then%>
								<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeDown"><%= round(objRSgua.Fields ("l5barps").Value,2) %></span>
							<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeEven"><%= round(objRSgua.Fields ("l5barps").Value,2) %></span>
							<%end if %>														
							</td>
