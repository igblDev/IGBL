							<td class="big" style="vertical-align:middle;text-align:center;vertical-align:bottom;width:20%;padding: inherit;"><img class="img-fluid center-block" src="<%= (objRSfor.Fields("image").Value)%>"></td>
							<td class="big" style="vertical-align:middle;text-align:left;width:34%;">
							<%if (len(objRSfor.Fields("firstName").Value) + len(objRSfor.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSfor.Fields("PID").Value %>">
									<%=left(objRSfor.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSfor.Fields("lastName").Value,14)%></a>
							<%else%>
								<a class="blue" href="playerprofile.asp?pid=<%=objRSfor.Fields("PID").Value %>">
									<%=objRSfor.Fields("firstName").Value%>&nbsp;<%=objRSfor.Fields("lastName").Value%></a>
							<%end if%>		
							<br><small><span class="greenTrade"><%= objRSfor.Fields("team")%></span> | <span class="orange"><%= objRSfor.Fields("pos")%></span>	
							<% if objRSfor.Fields("rentalPlayer").Value = true then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>
							<%if objRSfor.Fields("IR").Value = true then %>
								<strong>	<i class="fas fa-briefcase-medical red"></i></strong>
							<%end if%>	
							<% if objRSfor.Fields("pendingtrade").Value = true then %>
								<strong>	<i class="far fa-exchange"></i></strong>
							<%end if%>
							<% if objRSfor.Fields("pendingWaiver").Value = true then %>
								<strong>	<i class="fas fa-user-clock auctionText"></i></strong>
							<%end if%>
							<br><span class="gameTip"><%= objRSfor.Fields("opponent")%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%= wtime%></small>
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;width:18%;" class="text-center">
								<span class="badgeBlue"><%= round(objRSfor.Fields ("barps").Value,2) %></span>
							</td>
							<% if CDbl(objRSfor.Fields("l5barps").Value) > CDbl(objRSfor.Fields("barps").Value) then %>
								<td class="big" style="vertical-align:middle;text-align:center"><span class="badgeUp"><%= round(objRSfor.Fields ("l5barps").Value,2) %></span>
							<% elseif CDbl(objRSfor.Fields("barps").Value) > CDbl(objRSfor.Fields("l5barps").Value) then%>
								<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeDown"><%= round(objRSfor.Fields ("l5barps").Value,2) %></span>
							<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;width:18%;"><span class="badgeEven"><%= round(objRSfor.Fields ("l5barps").Value,2) %></span>
							<%end if %>														
							</td>
							<!--
							<td class="big" style="vertical-align:middle;text-align: center;width:9%;" class="text-center">	
								<span class="badgeUsage"><%= round(objRSfor.Fields ("usage").Value,0) %></span>												
							</td>
							-->