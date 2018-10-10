				<% 
					objRSLogs.close
					Set objRSLogs  = Server.CreateObject("ADODB.RecordSet")
					objRSLogs.Open  "SELECT * FROM qryResultHistoryPO " &_
					                "WHERE gamedate >= (select param_date from tblParameterCtl where param_name = 'PLAYOFF_START_DATE') " &_
									        "ORDER BY GameDate DESC" , objConn,3,3,1	
				%>
					<div class="panel panel-override">
						<div class="panel-body">
							<table class="table table-custom-black table-bordered table-striped table-condensed">
								<thead>
									<tr>
										<th>Play-Off Box Scores by Date</th>
									</tr>
								</thead>
								<tbody>
								<%
									While Not objRSLogs.eof
								%>
									<tr bgcolor="#FFFFFF">	
										<td style="vertical-align:middle;text-align:center;"><orangeIcon2><a class="blue" href="<%= objRSLogs.Fields ("formattedDate").Value %>.asp"><%= (FormatDateTime(objRSLogs.Fields ("GameDate").Value,1)) %></a></td>
									</tr>
								<%
									objRSLogs.MoveNext
									Wend
								%>
								</tbody>
							</table>
						</div>
					</div>