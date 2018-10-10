					<tr>
						<%if objRSNBASked.RecordCount > 0 then %>
							<td class ="big" style="vertical-align:middle;text-align:left;">
							<%if (len(objrsPlayerPos.Fields("firstName").Value) + len(objrsPlayerPos.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofileWork.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=left(objrsPlayerPos.Fields("firstName").Value,1)%>.&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%else%>
								<a class="blue" href="playerprofileWork.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=objrsPlayerPos.Fields("firstName").Value%>&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%end if%>					
							<br><small><span class="greenTrade"><%= objrsPlayerPos.Fields("team").Value%></span>&nbsp;|&nbsp;<span class="orange"><%=objrsPlayerPos.Fields("pos").Value%></span>
							<br><span class="gameTip"><%= objRSNBASked.Fields("opponent").Value %>&nbsp;<i class="far fa-clock"></i>&nbsp;<%=wtime%></small></span>
						<%else %>
							<td class ="big" style="vertical-align:middle;text-align:left;">					
							<%if (len(objrsPlayerPos.Fields("firstName").Value) + len(objrsPlayerPos.Fields("lastName").Value)) >= 17 then %>
								<a class="blue" href="playerprofileWork.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=left(objrsPlayerPos.Fields("firstName").Value,1)%>.&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%else%>
								<a class="blue" href="playerprofileWork.asp?pid=<%=objrsPlayerPos.Fields("PID").Value %>">
									<%=objrsPlayerPos.Fields("firstName").Value%>&nbsp;<%=objrsPlayerPos.Fields("lastName").Value%></a>
							<%end if%>
							<br><small><span class="greenTrade"><%= objrsPlayerPos.Fields("team").Value%></span>&nbsp;|&nbsp;<span class="orange"><%=objrsPlayerPos.Fields("pos").Value%></small></span>
						</td>
						<%end if%>		
						<td class="big" style="vertical-align: middle;"><%=round(objrsPlayerPos.Fields("barps").Value,2)%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("blk").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("ast").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("reb").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("ppg").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("stl").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("three").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("to").Value%></td>
						<td class="big" style="vertical-align: middle;"><%=objrsPlayerPos.Fields("numTdbls").Value%></td>
					</tr>	