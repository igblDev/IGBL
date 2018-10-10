							<td style="vertical-align:middle;text-align:center;width:15%;"><img class="img-responsive" src="<%= (objRSfor.Fields("image").Value)%>"></td>

							<td style="vertical-align:middle;text-align:left;width:35%;"><a class="blue" href="playerprofile.asp?pid=<%=objRSfor.Fields("PID").Value %>"><%=left(objRSfor.Fields("firstName").Value,1)%>.&nbsp;<%=left(objRSfor.Fields("lastName").Value,14)%></a>
	
							<br><orange><%= objRSfor.Fields("team")%> | <%= objRSfor.Fields("pos")%></orange>							
							<% if objRSfor.Fields("rentalPlayer").Value = true  then %>
							 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>
							<%end if%>							
							
							<br><gameTip><%= objRSfor.Fields("opponent")%>&nbsp;<span class="glyphicon glyphicon-time"></span> <%= wtime%> cst</gameTip>
							</td>							
							<td style="vertical-align:middle;" class="text-center">	
								<span class="badgeBlue"><%= round(objRSfor.Fields ("barps").Value,0) %></span>							
							</td>
							<% if CInt(objRSfor.Fields("l5barps").Value) > CInt(objRSfor.Fields("barps").Value) then %>
								<td style="vertical-align:middle;text-align:center"><span class="badgeUp"><%= round(objRSfor.Fields ("l5barps").Value,0) %></span></td>
							<% elseif CInt(objRSfor.Fields("barps").Value) > CInt(objRSfor.Fields("l5barps").Value) then%>
								<td style="vertical-align:middle;text-align:center"><span class="badgeDown"><%= round(objRSfor.Fields ("l5barps").Value,0) %></span></td>
							<%else%>
							<td style="vertical-align:middle;text-align:center"><span class="badgeEven"><%= round(objRSfor.Fields ("l5barps").Value,0) %></span></td>
							<%end if %>
							</td>
							<td style="vertical-align:middle;" class="text-center">	
								<span class="badgeUsage"><%= round(objRSfor.Fields ("usage").Value,0) %></span>															
							</td>