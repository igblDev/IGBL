</br>
<%
	While Not objRS.EOF
	
	tbonusHP1 = 0
	tbonusHP2 = 0		
	tbonusHP3 = 0
	tbonusHP4 = 0
	tbonusHP5 = 0		

	tbonusAP1 = 0
	tbonusAP2 = 0		
	tbonusAP3 = 0
	tbonusAP4 = 0
	tbonusAP5 = 0		
	
	penalty = 0 
	awayPen = 0 
	homePen = 0 
	
	'objRSAwaylogo.Open "SELECT * FROM TBLOWNERS WHERE ownerid = " & objRS.Fields("awayTeamOID").Value & " ",objConn,3,3,1 
	'objRSHomelogo.Open "SELECT * FROM TBLOWNERS WHERE ownerid = " & objRS.Fields("homeTeamOID").Value & " ",objConn,3,3,1 
	objRSAwaylogo.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & objRS.Fields("awayTeamOID").Value & "))", objConn
	objRSHomelogo.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & objRS.Fields("homeTeamOID").Value & "))", objConn

	if len(objRSAwaylogo.Fields("teamName")) >= 26 then 
		awayLogo = objRSAwaylogo.Fields("ShortName").value
	else
		awayLogo = objRSAwaylogo.Fields("teamName").value
	end if

	if len(objRSHomelogo.Fields("teamName")) >= 26 then
		homeLogo  = objRSHomelogo.Fields("ShortName").value 
	else
		homeLogo  = objRSHomelogo.Fields("teamName").value  
	end if

	
	<!--INITIALIZE 3DBL COUNTERS -->
	
	tbonusHP1 = 0
	tbonusHP2 = 0		
	tbonusHP3 = 0
	tbonusHP4 = 0
	tbonusHP5 = 0		

	tbonusAP1 = 0
	tbonusAP2 = 0		
	tbonusAP3 = 0
	tbonusAP4 = 0
	tbonusAP5 = 0		

	<!--BONUS INDICATOR CHECK FOR HOME PLAYERS -->

	if Cint(objRS.Fields("hp1BLK").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRS.Fields("hp1ASS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRS.Fields("hp1RBS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRS.Fields("hp1PTS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRS.Fields("hp1STL").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRS.Fields("hp13S").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	
	if Cint(objRS.Fields("hp2BLK").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp2ASS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp2RBS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp2PTS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp2STL").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp23S").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRS.Fields("hp3BLK").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp3ASS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp3RBS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp3PTS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp3STL").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp33S").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRS.Fields("hp4BLK").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRS.Fields("hp4ASS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRS.Fields("hp4RBS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRS.Fields("hp4PTS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRS.Fields("hp4STL").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRS.Fields("hp43S").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if			
	
	if Cint(objRS.Fields("hp5BLK").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRS.Fields("hp5ASS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRS.Fields("hp5RBS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRS.Fields("hp5PTS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRS.Fields("hp5STL").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRS.Fields("hp53S").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if				
	
	<!--BONUS INDICATOR CHECK FOR AWAY PLAYERS -->
	if Cint(objRS.Fields("ap1BLK").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRS.Fields("ap1ASS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRS.Fields("ap1RBS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRS.Fields("ap1PTS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRS.Fields("ap1STL").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRS.Fields("ap13S").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	
	if Cint(objRS.Fields("ap2BLK").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap2ASS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap2RBS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap2PTS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap2STL").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap23S").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRS.Fields("ap3BLK").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap3ASS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap3RBS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap3PTS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap3STL").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap33S").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRS.Fields("ap4BLK").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRS.Fields("ap4ASS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRS.Fields("ap4RBS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRS.Fields("ap4PTS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRS.Fields("ap4STL").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRS.Fields("ap43S").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if			
	
	if Cint(objRS.Fields("ap5BLK").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRS.Fields("ap5ASS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRS.Fields("ap5RBS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRS.Fields("ap5PTS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRS.Fields("ap5STL").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRS.Fields("ap53S").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if				

	'tbonusAP2 = 1
	'Response.Write "Value is: " & tbonusAP2 & ".<br>"
			
	penalty       = CInt(objRS.Fields("penOID").Value)
	if penalty    = CInt(objRS.Fields("homeTeamOID").Value) then
		homePen     = penalty
	elseif penalty= CInt(objRS.Fields("awayTeamOID").Value) then 
		awayPen     = penalty
	end if  
	'Response.Write "Value is: " & penalty & ".<br>"
	
	hscore = CDBL(objRS.Fields("htotBarps").Value)
	ascore = CDBL(objRS.Fields("atotBarps").Value)
	
	hp1Last =  left(objRS.Fields("hp1Last").Value,3)
	hp2Last =  left(objRS.Fields("hp2Last").Value,3)
	hp3Last =  left(objRS.Fields("hp3Last").Value,3)
	hp4Last =  left(objRS.Fields("hp4Last").Value,3)
	hp5Last =  left(objRS.Fields("hp5Last").Value,3)
	ap1Last =  left(objRS.Fields("ap1Last").Value,3)
	ap2Last =  left(objRS.Fields("ap2Last").Value,3)
	ap3Last =  left(objRS.Fields("ap3Last").Value,3)
	ap4Last =  left(objRS.Fields("ap4Last").Value,3)
	ap5Last =  left(objRS.Fields("ap5Last").Value,3)
	
	if hp1Last = "***" then
		hp1Last = false
	else
		hp1Last = objRS.Fields("hp1Last").Value
	end if

	if hp2Last = "***" then
		hp2Last = false
	else
		hp2Last = objRS.Fields("hp2Last").Value
	end if
	
	if hp3Last = "***" then
		hp3Last = false
	else
		hp3Last = objRS.Fields("hp3Last").Value
	end if

	if hp4Last = "***" then
		hp4Last = false
	else
		hp4Last = objRS.Fields("hp4Last").Value
	end if

	if hp5Last = "***" then
		hp5Last = false
	else
		hp5Last = objRS.Fields("hp5Last").Value
	end if
	
	if ap1Last = "***" then
		ap1Last = false
	else
		ap1Last = objRS.Fields("ap1Last").Value
	end if

	if ap2Last = "***" then
		ap2Last = false
	else
		ap2Last = objRS.Fields("ap2Last").Value
	end if 
	
	if ap3Last = "***" then
		ap3Last = false
	else
		ap3Last = objRS.Fields("ap3Last").Value
	end if
	
	if ap4Last = "***" then
		ap4Last = false
	else
		ap4Last = objRS.Fields("ap4Last").Value
	end if
	
	if ap5Last = "***" then
		ap5Last = false
	else
		ap5Last = objRS.Fields("ap5Last").Value
	end if
	'if (objRS.Fields("htotBarps").Value) = (objRS.Fields("atotBarps").Value) then
	if hscore = ascore then	
		hTBtos  = CInt(objRS.Fields("htotTOs").Value) 
		hTBstls = CInt(objRS.Fields("htotSTL").Value)
		hTBblks = CInt(objRS.Fields("htotBLK").Value)
		hTB3ptrs= CInt(objRS.Fields("htot3S").Value) 
		hTBrebs = CInt(objRS.Fields("htotRBS").Value)
		hTBass  = CInt(objRS.Fields("htotASS").Value)
		hTBpts  = CInt(objRS.Fields("htotPTS").Value)

		aTBtos  = CInt(objRS.Fields("atotTOs").Value) 
		aTBstls = CInt(objRS.Fields("atotSTL").Value)
		aTBblks = CInt(objRS.Fields("atotBLK").Value)
		aTB3ptrs= CInt(objRS.Fields("atot3S").Value) 
		aTBrebs = CInt(objRS.Fields("atotRBS").Value)
		aTBass  = CInt(objRS.Fields("atotASS").Value)
		aTBpts  = CInt(objRS.Fields("atotPTS").Value)			
	else 		

		hTBtos  = 0
		hTBstls = 0
		hTBblks = 0
		hTB3ptrs= 0
		hTBrebs = 0
		hTBass  = 0
		hTBpts  = 0

		aTBtos  = 0 
		aTBstls = 0
		aTBblks = 0
		aTB3ptrs= 0
		aTBrebs = 0
		aTBass  = 0
		aTBpts  = 0
	end if
	
	
	wTeamNameH = replace(UCase(objRSHomelogo.Fields("shortname").Value), "THE ", "")
	wTeamNameA  = replace(UCase(objRSAwaylogo.Fields("shortname").Value), "THE ", "")

%>
		
	<table class="table table-bordered table-condensed">
		<tr style="background-color:white;text-align:center;vertical-align:middle">
			<td style="width:50%"><img class="img-responsive" src="<%=objRSHomelogo.Fields("teamlogo").Value%>" style="margin: 0 auto;width:auto;height:100px;"></td>
			<td style="width:50%"><img class="img-responsive" src="<%=objRSAwaylogo.Fields("teamlogo").Value%>" style="margin: 0 auto;width:auto;height:100px;"></td>
		</tr>
	</table>
	<table class="table table-responsive table-bordered table-custom-black table-condensed">
		<tr style="background-color: #ddd;">
			<!--Checks to verify if game is tied. If so implement tie-breaking rules listed above -->
			<%if hscore <> ascore then %>
				<% if hscore > ascore then%>
					<% if homePen > 0 then %>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><span class="badgeHomeTeam big" style="border-radius: 16px;"><% =(objRS.Fields("htotBarps").Value)%></span></span>&nbsp;<redW>P</redW>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><span class="badgeHomeTeam big" style="border-radius: 16px;"><% =(objRS.Fields("htotBarps").Value)%></span></span>
						</th>
					<%end if%>
				<%else%>
					<th class="big" style="font-size:12px;width:50%">
						<span class="big"><%=wTeamNameH%></span><br>
						<span class="big" style="font-size:15px;"><span class="badgeHomeTeam big" style="border-radius: 16px;"><% =(objRS.Fields("htotBarps").Value)%></span></span>
					</th>
				<%end if %>
				<%if ascore > hscore then %>
					<% if awayPen > 0 then %>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><span class="badgeAwayTeam big" style="border-radius: 16px;"><% =(objRS.Fields("atotBarps").Value)%></span></span>&nbsp;<redW>P</redW>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><span class="badgeAwayTeam big" style="border-radius: 16px;"><% =(objRS.Fields("atotBarps").Value)%></span></span>
						</th>
					<%end if%>
				<%else%>
					<th class="big" style="font-size:12px;width:50%">
						<span class="big"><%=wTeamNameA%></span><br>
						<span class="big" style="font-size:15px;"><span class="badgeAwayTeam big" style="border-radius: 16px;"><% =(objRS.Fields("atotBarps").Value)%></span></span>
					</th>
				<%end if %>
			<%elseif hTBtos <> aTBtos then %>
				<!--TO TIE-BREAKER -->		
					<% if hTBtos < aTBtos then %>
						<%htToWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
								<span class="big" style="font-size:15px;"><span class="badgeHomeTeam big" style="border-radius: 16px;"><% =(objRS.Fields("htotBarps").Value)%></span></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><span class="badgeHomeTeam big" style="border-radius: 16px;"><% =(objRS.Fields("htotBarps").Value)%></span></span>
						</th>
					<%end if%>
					<% if aTBtos < hTBtos then %>
						<%atToWon = true %>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
						<span class="big" style="font-size:15px;"><span class="badgeAwayTeam big" style="border-radius: 16px;"><% =(objRS.Fields("atotBarps").Value)%></span></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
						<span class="big" style="font-size:15px;"><span class="badgeAwayTeam big" style="border-radius: 16px;"><% =(objRS.Fields("atotBarps").Value)%></span></span>
						</th>
					<%end if%>
				<!--END OF 	TO TIE-BREAKER -->	
			<%elseif hTBstls <> aTBstls then%>
				<!--STEAL TIE-BREAKER -->			
					<% if hTBstls > aTBstls then%>
						<%htStlWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%end if%>
					<% if aTBstls > hTBstls then%>
					<%atStlWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%end if%>
					<!--END OF STEAL TIE-BREAKER -->	
			<%elseif hTBblks <> aTBblks then%>	
				<!--BLOCK TIE-BREAKER -->
					<% if hTBblks > aTBblks then%>
						<%htBlkWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%end if%>
					<% if aTBblks > hTBblks then%>
						<%atBlkWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%end if%>
				<!--END OF BLOCK TIE-BREAKER -->
			<%elseif hTB3ptrs <> aTB3ptrs then%>	
				<!--BLOCK 3s-BREAKER -->
					<% if hTB3ptrs > aTB3ptrs then%>
						<%ht3Won = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("htotBarps").Value)%></span>
						</th>
					<%end if%>
					<% if aTB3ptrs > hTB3ptrs then%>
						<%at3Won = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
							<span class="big" style="font-size:15px;"><% =(objRS.Fields("atotBarps").Value)%></span>
						</th>
					<%end if%>
				<!--END OF BLOCK TIE-BREAKER -->
			<% end if %>
		</tr>
	</table>
	</br>			
	<ul class="nav nav-tabs">
		<% if (hscore > ascore) or (htToWon = true) or (htStlWon = true) or (ht3Won = true) then%>
			<li class="active big">
				<a data-toggle="tab" href="#<%=hcount%>"><%=wTeamNameH%>
					<span style="color:yellowgreen;font-weight:bold;"<i class="fas fa-check-circle fa-lg"></i></span>
				</a>
			</li>  
		<% else %>	
			<li class="big"><a data-toggle="tab" href="#<%=hcount%>"><%=wTeamNameH%></a></li>
		<% end if %>
		<% if (ascore > hscore) or (atToWon = true) or (atStlWon = true) or (at3Won = true) then %>
			<li class="active big">
				<a data-toggle="tab" href="#<%=acount%>"><%=wTeamNameA%>
					<span style="color:darkorange;font-weight:bold;"><i class="fas fa-check-circle fa-lg"></i></span>
				</a>
			</li>  
		<%else %>
			<li class="big"><a data-toggle="tab" href="#<%=acount%>"><%=wTeamNameA%></a></li>  
		<%end if %>		
	</ul> 
	<div class="tab-content">	
		<div class="div.contained_table">
      <table class="table table-responsive  table-condensed table-bordered">       
        <thead>
					<tr style="background-color:#dddddd;font-weight:bold;color:grey;">
						<td style="vertical-align:middle;text-align:center;width:20%;"><span style="color:black;">PLAYER</span></td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">M</span>ins</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">B</span>lks</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">A</span>sts</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">R</span>ebs</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">P</span>ts</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">S</span>tls</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">3</span>s</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">T</span>os</td>	
						<th class="big" width="10%"><span style="color:black;"><i class="fas fa-basketball-ball"></i></th>						
					</tr>
        </thead>
				 <tr style="font-weight:normal;background-color:white;">	
					<% if hp1Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(hp1Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("hp1MP").Value)%></td>
						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp1BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp1BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp1BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp1ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp1ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp1ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp1RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp1RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp1RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp1PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp1PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp1PTS").Value)%></td>
						<%end if%>

						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp1STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp1STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp1STL").Value)%></td>
						<%end if%>						

						<% if tbonushp1 >= 3 and CInt(objRS.Fields("hp13S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRS.Fields("hp13S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp13S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("hp1TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("hp1Barps").Value)%></td>
					</tr>

				 <tr style="font-weight:normal;background-color:white;">	
					<% if hp2Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(hp2Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("hp2MP").Value)%></td>	
						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp2BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp2BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp2BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp2ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp2ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp2ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp2RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp2RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp2RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp2PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp2PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp2PTS").Value)%></td>
						<%end if%>

						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp2STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp2STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp2STL").Value)%></td>
						<%end if%>						

						<% if tbonushp2 >= 3 and CInt(objRS.Fields("hp23S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp23S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp23S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("hp2TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("hp2Barps").Value)%></td>
					</tr>
					<tr style="font-weight:normal;background-color:white;">	
					<% if hp3Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(hp3Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("hp3MP").Value)%></td>
						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp3BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp3BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp3BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp3ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp3ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp3ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp3RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp3RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp3RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp3PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp3PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp3PTS").Value)%></td>
						<%end if%>

						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp3STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp3STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp3STL").Value)%></td>
						<%end if%>						

						<% if tbonushp3 >= 3 and CInt(objRS.Fields("hp33S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp33S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp33S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("hp3TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("hp3Barps").Value)%></td>
					</tr>
				  <tr style="font-weight:normal;background-color:white;">	
					<% if hp4Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(hp4Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("hp4MP").Value)%></td>
						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp4BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp4BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp4BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp4ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp4ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp4ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp4RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp4RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp4RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp4PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp4PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp4PTS").Value)%></td>
						<%end if%>

						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp4STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp4STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp4STL").Value)%></td>
						<%end if%>						

						<% if tbonushp4 >= 3 and CInt(objRS.Fields("hp43S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp43S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp43S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("hp4TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("hp4Barps").Value)%></td>
					</tr>
				  <tr style="font-weight:normal;background-color:white;">		
					<% if hp5Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(hp5Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("hp5MP").Value)%></td>
						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp5BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp5BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp5BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp5ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp5ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp5ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp5RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp5RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp5RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp5PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp5PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp5PTS").Value)%></td>
						<%end if%>

						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp5STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp5STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp5STL").Value)%></td>
						<%end if%>						

						<% if tbonushp5 >= 3 and CInt(objRS.Fields("hp53S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("hp53S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("hp53S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("hp5TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("hp5Barps").Value)%></td>
					</tr>
			  <tr style="text-align:center;vertical-align:middle;font-size:12px;font-weight:700;background-color:black;color:yellowgreen;">						
					<% if homePen > 0 then %>
          <td class="big" style="vertical-align:middle;" width="20%">PENALTY</td>
					<%else%>
          <td class="big" style="vertical-align:middle;" width="20%">TEAM</td>
					<%end if%>	
					<td style="vertical-align:middle;"><i class="fa fa-clock-o" aria-hidden="true"></i></td>					
					<%if htBlkWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("htotBLK").Value)%></span></td>	
					<%else%>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotBLK").Value)%></td>
					<%end if%>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotASS").Value)%></td>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotRBS").Value)%></td>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotPTS").Value)%></td>
					<%if htStlWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("htotSTL").Value)%></span></td>
					<%else%>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotSTL").Value)%></td>
					<%end if %>
					<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htot3S").Value)%></td>
					<% if htToWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("htotTOs").Value)%></span></td>
					<%else %>
          <td class="big" style="vertical-align:middle;"><%=(objRS.Fields("htotTOs").Value)%></td>
					<%end if %>
					<td style="vertical-align:middle;text-align:center;background-color:yellowgreen;"><span class="badgeHomeTeam"><% =(objRS.Fields("htotBarps").Value)%></span></td>
        </tr>          
      </table>
	  </div>
		<div class="div.contained_table">
			<table class="table table-responsive  table-condensed table-bordered">    
				<thead>
					<tr style="background-color:#dddddd;font-weight:bold;color:grey;">
						<td style="vertical-align:middle;text-align:center;width:20%;"><span style="color:black;">PLAYER</span></td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">M</span>ins</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">B</span>lks</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">A</span>sts</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">R</span>ebs</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">P</span>ts</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">S</span>tls</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">3</span>s</td>
						<td style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">T</span>os</td>	
						<th class="big" width="10%"><span style="color:black;"><i class="fas fa-basketball-ball"></i></th>						
					</tr>
				</thead>
				 <tr style="font-weight:normal;background-color:white;">	
					<% if ap1Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(ap1Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("ap1MP").Value)%></td>
						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap1BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap1BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap1BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap1ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap1ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap1ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap1RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap1RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap1RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap1PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap1PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap1PTS").Value)%></td>
						<%end if%>

						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap1STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap1STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap1STL").Value)%></td>
						<%end if%>						

						<% if tbonusap1 >= 3 and CInt(objRS.Fields("ap13S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap13S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap13S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("ap1TOs").Value)%></td>
						<td class="big" ><span ><%=(objRS.Fields("ap1Barps").Value)%></span></td>
					</tr>
				 <tr style="font-weight:normal;background-color:white;">	
					<% if ap2Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(ap2Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("ap2MP").Value)%></td>
						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap2BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap2BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap2BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap2ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap2ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap2ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap2RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap2RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap2RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap2PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap2PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap2PTS").Value)%></td>
						<%end if%>

						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap2STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap2STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap2STL").Value)%></td>
						<%end if%>						

						<% if tbonusAP2 >= 3 and CInt(objRS.Fields("ap23S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap23S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap23S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("ap2TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("ap2Barps").Value)%></td>
					</tr>
				  <tr style="font-weight:normal;background-color:white;">	
					<% if ap3Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(ap3Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("ap3MP").Value)%></td>
						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap3BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap3BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap3BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap3ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap3ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap3ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap3RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap3RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap3RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap3PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap3PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap3PTS").Value)%></td>
						<%end if%>

						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap3STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap3STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap3STL").Value)%></td>
						<%end if%>						

						<% if tbonusap3 >= 3 and CInt(objRS.Fields("ap33S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap33S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap33S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("ap3TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("ap3Barps").Value)%></td>
					</tr>
				 <tr style="font-weight:normal;background-color:white;">	
					<% if ap4Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(ap4Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("ap4MP").Value)%></td>
						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap4BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap4BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap4BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap4ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap4ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap4ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap4RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap4RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap4RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap4PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap4PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap4PTS").Value)%></td>
						<%end if%>

						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap4STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap4STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap4STL").Value)%></td>
						<%end if%>						

						<% if tbonusap4 >= 3 and CInt(objRS.Fields("ap43S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap43S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap43S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("ap4TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("ap4Barps").Value)%></td>
					</tr>
				  <tr style="font-weight:normal;background-color:white;">	
					<% if ap5Last = false then %>
						<td class="opp big" style="text-align:left">
							<span class="fa-stack fa-xs">
							<i class="fa fa-user fa-stack-1x"></i>
							<i class="fa fa-ban fa-stack-2x text-danger"></i>
							</span>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;width:20%;"><%=left(ap5Last,10)%></td>
					<%end if %>
						<td><%=(objRS.Fields("ap5MP").Value)%></td>
						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap5BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap5BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap5BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap5ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap5ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap5ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap5RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap5RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap5RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap5PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap5PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap5PTS").Value)%></td>
						<%end if%>

						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap5STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap5STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap5STL").Value)%></td>
						<%end if%>						

						<% if tbonusap5 >= 3 and CInt(objRS.Fields("ap53S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRS.Fields("ap53S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRS.Fields("ap53S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRS.Fields("ap5TOs").Value)%></td>
						<td class="big" ><%=(objRS.Fields("ap5Barps").Value)%></td>
				</tr>
			  <tr style="text-align:center;vertical-align:middle;font-size:12px;font-weight:700;background-color:black;color:darkorange;">										
						<% if awayPen > 0 then %>
						<td class="big" style="vertical-align:middle;" width="20%">PENALTY</td>
						<%else%>
						<td class="big" style="vertical-align:middle;" width="20%">TEAM</td>
						<%end if%>
							<td style="vertical-align:middle;"><i class="fa fa-clock-o" aria-hidden="true"></i></td>		
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotBLK").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotASS").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotRBS").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotPTS").Value)%></td>							
							<%if atStlWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("atotSTL").Value)%></span></td>
							<%else%>
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotSTL").Value)%></td>
							<%end if%>
							<%if atBlkWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("atot3S").Value)%></span></td>		
							<%else%>	
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atot3S").Value)%></td>	
							<%end if%>
							<% if atToWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRS.Fields("atotTOs").Value)%></span></td>
							<%else%>
							<td class="big" style="vertical-align:middle;"><%=(objRS.Fields("atotTOs").Value)%></td>
							<%end if%>
							<td style="background-color:darkorange;" ><span class="badgeAwayTeam"><% =(objRS.Fields("atotBarps").Value)%></span></td>
				</tr>
			</table>	
		</div>	
	</div>
</br>	
