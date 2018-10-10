							<td class="big" style="vertical-align:middle;text-align:center;vertical-align:bottom;width:20%;padding: inherit;"><img class="img-fluid center-block" src="<%= (objRScen.Fields("image").Value)%>"></td>
							<td class="big" style="vertical-align:middle;text-align:left;width:34%;">
							<%if (len(objRScen.Fields("firstName").Value) + len(objRScen.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objRScen.Fields("PID").Value %>">
									<%=left(objRScen.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRScen.Fields("lastName").Value,14)%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objRScen.Fields("PID").Value %>">
									<%=objRScen.Fields("firstName").Value%>&nbsp;<%=objRScen.Fields("lastName").Value%></a>
							<%end if%>		
							<br><small><span class="greenTrade"><%= objRScen.Fields("team")%></span> | <span class="orange"><%= objRScen.Fields("pos")%></span>	
							<% if objRScen.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							<%if objRScen.Fields("IR").Value = true then %>
								<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
							<%end if%>	
							<% if objRScen.Fields("pendingtrade").Value = true then %>
								<strong>	<i class="far fa-exchange"></i></strong>
							<%end if%>
							<% if objRScen.Fields("pendingWaiver").Value = true then %>
								<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
							<%end if%>
							<br><span class="gameTip"><%= objRScen.Fields("opponent")%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%= wtime%></small>
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;width:18%;" class="text-center">
								<span class="badgeBlue"><%= round(objRScen.Fields ("barps").Value,2) %></span>
							</td>
							<% if CDbl(objRScen.Fields("l5barps").Value) > CDbl(objRScen.Fields("barps").Value) then %>
								<td class="big" style="vertical-align:middle;text-align:center"><span class="badgeUp"><%= round(objRScen.Fields ("l5barps").Value,2) %></span>
							<% elseif CDbl(objRScen.Fields("barps").Value) > CDbl(objRScen.Fields("l5barps").Value) then%>
								<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeDown"><%= round(objRScen.Fields ("l5barps").Value,2) %></span>
							<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeEven"><%= round(objRScen.Fields ("l5barps").Value,2) %></span>
							<%end if %>														
							</td>
