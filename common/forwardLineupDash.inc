							<td class="big" style="vertical-align:middle;text-align:center;vertical-align:bottom;width:20%;padding: inherit;"><img class="img-responsive center-block" src="<%= (objRSfor.Fields("image").Value)%>"></td>
							<td class="big" style="vertical-align:middle;text-align:left;">
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
							<br><span class="gameTip"><%= objRSfor.Fields("opponent")%>&nbsp;<i class="fas fa-clock"></i>&nbsp;<%= wtime%></small>
							</td>
							<td class="big" style="vertical-align:middle;text-align: center;" class="text-center">
								<%= round(objRSfor.Fields ("barps").Value,2) %>
							</td>
							<% if CInt(objRSfor.Fields("l5barps").Value) > CInt(objRSfor.Fields("barps").Value) then %>
								<td class="big" style="vertical-align:middle;text-align:center"><%= round(objRSfor.Fields ("l5barps").Value,2) %>
							<% elseif CInt(objRSfor.Fields("barps").Value) > CInt(objRSfor.Fields("l5barps").Value) then%>
								<td class="big" style="vertical-align:middle;text-align:center;"><%= round(objRSfor.Fields ("l5barps").Value,2) %>
							<%else%>
							<td class="big" style="vertical-align:middle;text-align:center;"><%= round(objRSfor.Fields ("l5barps").Value,2) %>
							<%end if %>														
							</td>
