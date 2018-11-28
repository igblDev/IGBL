<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-analysis">
				<i class="fa fa-exclamation-triangle fa-lg red" aria-hidden="true"></i> Trade Analysis Summary<br><br>
				 <table class="table table-custom table-responsive table-bordered table-condensed">
						<tr>
							<th colspan="2" style="color:black;width:50%;font-size:12px !important;">Outgoing</th>
							<th colspan="2" style="color:black;width:50%;font-size:12px !important;">Incoming</th>
						</tr>
						<tr style="background-color:white;text-align:left;font-weight:bold;">
							<td class="big red" style="width:32%;text-align:left;vertical-align:top !important;"><%= saPlayer1Name %></td>		
							<td class="big" style="color:black;text-align:left;width:18%;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=saRegGameCnt1%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=saPOGameCnt1%></td>		
							<td class="big greenIcon" style="color:#468847;width:32%;text-align:left;vertical-align:top !important;"><%= stPlayer1Name %></td>							
							<td class="big" style="color:black;text-align:left;width:18%;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=stRegGameCnt1%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=stPOGameCnt1%></td>						
						</tr>
						<%		
						if saPlayer2Name <> "" or stPlayer2Name <> "" then
						%>		
						<tr style="background-color:white;text-align:left;font-weight:bold;">
							<td class="big" style="color:#9a1400;text-align:left;vertical-align:top !important;"><%= saPlayer2Name %> </td>
							<td class="big" style="color:black;text-align:left;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=saRegGameCnt2%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=saPOGameCnt2%></td>		
							<td class="big" style="color:#468847;text-align:left;vertical-align:top !important;"><%= stPlayer2Name %></td>								
							<td class="big" style="color:black;text-align:left;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=stRegGameCnt2%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=stPOGameCnt2%></td>
						 </tr>
						<%		
						end if
						if saPlayer3Name <> "" or stPlayer3Name <> "" then
						%>		
						<tr style="background-color:white;text-align:left;font-weight:bold;">
							 <td class="big" style="color:#9a1400;text-align:left;vertical-align:top !important;"><%= saPlayer3Name%></td>		
							 <td class="big" style="color:black;text-align:left;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=saRegGameCnt3%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=saPOGameCnt3%></td>		
							 <td class="big" style="color:#468847;width:30%;text-align:left;vertical-align:top !important;"><%= stPlayer3Name %></td>							
							 <td class="big" style="color:black;text-align:left;"><span style="color:#01579B;font-weight:bold;">REG:</span> <%=stRegGameCnt3%><br><span style="color:#01579B;font-weight:bold;">PO:</span> <%=stPOGameCnt3%></td>
						 </tr>
						<%		
						end if
						%>								
					</table>
					<br>
					<table class="table table-custom table-responsive table-bordered table-condensed">
						<tr style="text-align:center;"	>
						 <th colspan="2" style="text-align:center;color:black;width:32%;font-size:12px !important;">Positive</th>
						 <th colspan="2" style="text-align:center;color:black;width:32%;font-size:12px !important;">Negative</th>
						 <th colspan="2" style="text-align:center;color:black;width:36%;font-size:12px !important;">Neutral [+/- <%=wNeutralVal%>]</th>
						</tr>
						<tr style="background-color:white;color:black;text-align:center;">
							<td class="big" style="color:#468847;width:16%;vertical-align: middle;"><greenIcon><i class="fa fa-thumbs-up" aria-hidden="true"></i></greenIcon></td>							
							<td class="big" style="color:black;width:16%;"><%= wPositive %></td>							
							<td class="big" style="color:#9a1400;width:16%;vertical-align: middle;"><redIcon><i class="fa fa-thumbs-down" aria-hidden="true"></i></redIcon></td>							
							<td class="big" style="color:black;width:16%;"><%= wNegative %></td>							
							<td class="big" style="color:gold;width:18%;vertical-align: middle;"><evenIcon><i class="fa fa-balance-scale" aria-hidden="true"></i></evenIcon></td>							
							<td class="big" style="color:black;width:18%;"><%= wEven %></td>
						</tr>
						</table>
						</br>
						<!--<table class="table table-custom table-responsive table-bordered table-condensed">
						<tr>
						 <th colspan="6" style="color:black;width:32%;font-size:12px !important;">Re-Forecast Trade Analysis Summary</th>
						</tr>
						<tr style="background-color:white;color:black;">
							<td class="big" colspan="2">
								<select class="form-control input-sm" name="reForcast">
									<option value="">Select Neutral Factor Indicator</option>
									<option value="0">0 - Neutral Factor</option>
									<option value="1">1 - Neutral Factor</option>
									<option value="2">2 - Neutral Factor</option>
									<option value="3">3 - Neutral Factor</option>
									<option value="4">4 - Neutral Factor</option>			 
									<option value="5">5 - Neutral Factor</option>
									<option value="6">6 - Neutral Factor</option>
									<option value="7">7 - Neutral Factor</option>
									<option value="8">8 - Neutral Factor</option>		
									<option value="9">9 - Neutral Factor</option>
									<option value="10">10 - Neutral Factor</option>
									<option value="11">11 - Neutral Factor</option>
									<option value="12">12 - Neutral Factor</option>			 
									<option value="13">13 - Neutral Factor</option>
									<option value="14">14 - Neutral Factor</option>
									<option value="15">15 - Neutral Factor</option>
									</select>
							</td>
							<td class="big" colspan="4">
								<button type="submit" value="Reforecast" name="Action" class="btn btn-block btn-default btn-md"><span class="glyphicon glyphicon-save"></span>&nbsp;Re-Forecast Trade</button>
							</td>
						</tr>
					</table>-->
			</div>
		</div>
	</div>
</div>	