							<td class="big" style="vertical-align:middle;text-align:center;width:10%;">
								<%= objrsStarters.Fields("pos")%>&nbsp;
							</td>
							<td class="big" style="vertical-align:middle;text-align:left;width:60%;">
								<a class="blue" href="playerprofile.asp?pid=<%=objrsStarters.Fields("PID").Value %>">
								<%=objrsStarters.Fields("firstName").Value%>&nbsp;<%=objrsStarters.Fields("lastName").Value%></a>&nbsp;
								<%=objrsStarters.Fields("team")%>&nbsp;
								<% if objrsStarters.Fields("rentalPlayer").Value = true then %>
								 <strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>&nbsp;
								<%end if%>
							</td>
							<td class="big" style="vertical-align:middle;text-align:center;width:20%;">
								<%= objrsStarters.Fields("opponent")%>
							</td>
						
							