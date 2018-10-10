<%
	
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
	
	objRSAwaylogo.Close
	objRSHomelogo.Close

	objRSAwaylogo.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & objRBox.Fields("awayTeamOID").Value & "))", objConn
	objRSHomelogo.Open "SELECT * FROM qryTeams WHERE (((qryTeams.OwnerID)=" & objRBox.Fields("homeTeamOID").Value & "))", objConn

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

	if Cint(objRBox.Fields("hp1BLK").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRBox.Fields("hp1ASS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRBox.Fields("hp1RBS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRBox.Fields("hp1PTS").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRBox.Fields("hp1STL").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	if Cint(objRBox.Fields("hp13S").Value) >= 10 then 
		tbonushp1 = tbonushp1 + 1
	end if

	
	if Cint(objRBox.Fields("hp2BLK").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp2ASS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp2RBS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp2PTS").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp2STL").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp23S").Value) >= 10 then 
		tbonushp2 = tbonushp2 + 1
	end if

	if Cint(objRBox.Fields("hp3BLK").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp3ASS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp3RBS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp3PTS").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp3STL").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp33S").Value) >= 10 then 
		tbonushp3 = tbonushp3 + 1
	end if

	if Cint(objRBox.Fields("hp4BLK").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRBox.Fields("hp4ASS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRBox.Fields("hp4RBS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRBox.Fields("hp4PTS").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRBox.Fields("hp4STL").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if

	if Cint(objRBox.Fields("hp43S").Value) >= 10 then 
		tbonushp4 = tbonushp4 + 1
	end if			
	
	if Cint(objRBox.Fields("hp5BLK").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRBox.Fields("hp5ASS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRBox.Fields("hp5RBS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRBox.Fields("hp5PTS").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRBox.Fields("hp5STL").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if

	if Cint(objRBox.Fields("hp53S").Value) >= 10 then 
		tbonushp5 = tbonushp5 + 1
	end if				
	
	<!--BONUS INDICATOR CHECK FOR AWAY PLAYERS -->
	if Cint(objRBox.Fields("ap1BLK").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRBox.Fields("ap1ASS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRBox.Fields("ap1RBS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRBox.Fields("ap1PTS").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRBox.Fields("ap1STL").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	if Cint(objRBox.Fields("ap13S").Value) >= 10 then 
		tbonusap1 = tbonusap1 + 1
	end if

	
	if Cint(objRBox.Fields("ap2BLK").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap2ASS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap2RBS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap2PTS").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap2STL").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap23S").Value) >= 10 then 
		tbonusAP2 = tbonusAP2 + 1
	end if

	if Cint(objRBox.Fields("ap3BLK").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap3ASS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap3RBS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap3PTS").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap3STL").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap33S").Value) >= 10 then 
		tbonusap3 = tbonusap3 + 1
	end if

	if Cint(objRBox.Fields("ap4BLK").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRBox.Fields("ap4ASS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRBox.Fields("ap4RBS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRBox.Fields("ap4PTS").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRBox.Fields("ap4STL").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if

	if Cint(objRBox.Fields("ap43S").Value) >= 10 then 
		tbonusap4 = tbonusap4 + 1
	end if			
	
	if Cint(objRBox.Fields("ap5BLK").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRBox.Fields("ap5ASS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRBox.Fields("ap5RBS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRBox.Fields("ap5PTS").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRBox.Fields("ap5STL").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if

	if Cint(objRBox.Fields("ap53S").Value) >= 10 then 
		tbonusap5 = tbonusap5 + 1
	end if				

		
	penalty       = CInt(objRBox.Fields("penOID").Value)
	if penalty    = CInt(objRBox.Fields("homeTeamOID").Value) then
		homePen     = penalty
	elseif penalty= CInt(objRBox.Fields("awayTeamOID").Value) then 
		awayPen     = penalty
	end if  
	
	hscore = CDBL(objRBox.Fields("htotBarps").Value)
	ascore = CDBL(objRBox.Fields("atotBarps").Value)
	
	hp1Last =  left(objRBox.Fields("hp1Last").Value,3)
	hp2Last =  left(objRBox.Fields("hp2Last").Value,3)
	hp3Last =  left(objRBox.Fields("hp3Last").Value,3)
	hp4Last =  left(objRBox.Fields("hp4Last").Value,3)
	hp5Last =  left(objRBox.Fields("hp5Last").Value,3)
	ap1Last =  left(objRBox.Fields("ap1Last").Value,3)
	ap2Last =  left(objRBox.Fields("ap2Last").Value,3)
	ap3Last =  left(objRBox.Fields("ap3Last").Value,3)
	ap4Last =  left(objRBox.Fields("ap4Last").Value,3)
	ap5Last =  left(objRBox.Fields("ap5Last").Value,3)


	
	hp1PID  = objRBox.Fields("hp1PID").Value
	hp2PID  = objRBox.Fields("hp2PID").Value
	hp3PID  = objRBox.Fields("hp3PID").Value
	hp4PID  = objRBox.Fields("hp4PID").Value
	hp5PID  = objRBox.Fields("hp5PID").Value
	ap1PID  = objRBox.Fields("ap1PID").Value
	ap2PID  = objRBox.Fields("ap2PID").Value
	ap3PID  = objRBox.Fields("ap3PID").Value
	ap4PID  = objRBox.Fields("ap4PID").Value
	ap5PID  = objRBox.Fields("ap5PID").Value
	
	if hp1Last = "***" then
		hp1Last = false
	else
		hp1Last = objRBox.Fields("hp1Last").Value
	end if

	if hp2Last = "***" then
		hp2Last = false
	else
		hp2Last = objRBox.Fields("hp2Last").Value
	end if
	
	if hp3Last = "***" then
		hp3Last = false
	else
		hp3Last = objRBox.Fields("hp3Last").Value
	end if

	if hp4Last = "***" then
		hp4Last = false
	else
		hp4Last = objRBox.Fields("hp4Last").Value
	end if

	if hp5Last = "***" then
		hp5Last = false
	else
		hp5Last = objRBox.Fields("hp5Last").Value
	end if
	
	if ap1Last = "***" then
		ap1Last = false
	else
		ap1Last = objRBox.Fields("ap1Last").Value
	end if

	if ap2Last = "***" then
		ap2Last = false
	else
		ap2Last = objRBox.Fields("ap2Last").Value
	end if 
	
	if ap3Last = "***" then
		ap3Last = false
	else
		ap3Last = objRBox.Fields("ap3Last").Value
	end if
	
	if ap4Last = "***" then
		ap4Last = false
	else
		ap4Last = objRBox.Fields("ap4Last").Value
	end if
	
	if ap5Last = "***" then
		ap5Last = false
	else
		ap5Last = objRBox.Fields("ap5Last").Value
	end if
	'if (objRBox.Fields("htotBarps").Value) = (objRBox.Fields("atotBarps").Value) then
	if hscore = ascore then	
		hTBtos  = CInt(objRBox.Fields("htotTOs").Value) 
		hTBstls = CInt(objRBox.Fields("htotSTL").Value)
		hTBblks = CInt(objRBox.Fields("htotBLK").Value)
		hTB3ptrs= CInt(objRBox.Fields("htot3S").Value) 
		hTBrebs = CInt(objRBox.Fields("htotRBS").Value)
		hTBass  = CInt(objRBox.Fields("htotASS").Value)
		hTBpts  = CInt(objRBox.Fields("htotPTS").Value)

		aTBtos  = CInt(objRBox.Fields("atotTOs").Value) 
		aTBstls = CInt(objRBox.Fields("atotSTL").Value)
		aTBblks = CInt(objRBox.Fields("atotBLK").Value)
		aTB3ptrs= CInt(objRBox.Fields("atot3S").Value) 
		aTBrebs = CInt(objRBox.Fields("atotRBS").Value)
		aTBass  = CInt(objRBox.Fields("atotASS").Value)
		aTBpts  = CInt(objRBox.Fields("atotPTS").Value)			
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
	htToWon  = false
	atToWon  = false
	htStlWon = false
	atStlWon = false
	htBlkWon = false
	atBlkWon = false
	ht3Won   = false
	at3Won   = false
	
	wTeamNameH = replace(UCase(objRSHomelogo.Fields("shortname").Value), "THE ", "")
	wTeamNameA  = replace(UCase(objRSAwaylogo.Fields("shortname").Value), "THE ", "")

%>
	<table class="table table-bordered table-custom-black table-condensed">
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
						
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
						
						</th>
					<%end if%>
				<%else%>
					<th class="big" style="font-size:12px;width:50%">
						<span class="big"><%=wTeamNameH%></span><br>
					
					</th>
				<%end if %>
				<%if ascore > hscore then %>
					<% if awayPen > 0 then %>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
				
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
						
						</th>
					<%end if%>
				<%else%>
					<th class="big" style="font-size:12px;width:50%">
						<span class="big"><%=wTeamNameA%></span><br>
				
					</th>
				<%end if %>
			<%elseif hTBtos <> aTBtos then %>
				<!--TO TIE-BREAKER -->		
					<% if hTBtos < aTBtos then %>
						<%htToWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
							
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
						
						</th>
					<%end if%>
					<% if aTBtos < hTBtos then %>
						<%atToWon = true %>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
					
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
					
						</th>
					<%end if%>
				<!--END OF 	TO TIE-BREAKER -->	
			<%elseif hTBstls <> aTBstls then%>
				<!--STEAL TIE-BREAKER -->			
					<% if hTBstls > aTBstls then%>
						<%htStlWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
				
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
						
						</th>
					<%end if%>
					<% if aTBstls > hTBstls then%>
					<%atStlWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
					
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
						
						</th>
					<%end if%>
					<!--END OF STEAL TIE-BREAKER -->	
			<%elseif hTBblks <> aTBblks then%>	
				<!--BLOCK TIE-BREAKER -->
					<% if hTBblks > aTBblks then%>
						<%htBlkWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
				
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
			
						</th>
					<%end if%>
					<% if aTBblks > hTBblks then%>
						<%atBlkWon = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
			
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
			
						</th>
					<%end if%>
				<!--END OF BLOCK TIE-BREAKER -->
			<%elseif hTB3ptrs <> aTB3ptrs then%>	
				<!--BLOCK 3s-BREAKER -->
					<% if hTB3ptrs > aTB3ptrs then%>
						<%ht3Won = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
					
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameH%></span><br>
			
						</th>
					<%end if%>
					<% if aTB3ptrs > hTB3ptrs then%>
						<%at3Won = true%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
				
						</th>
					<%else%>
						<th class="big" style="font-size:12px;width:50%">
							<span class="big"><%=wTeamNameA%></span><br>
								</th>
					<%end if%>
				<!--END OF BLOCK TIE-BREAKER -->
			<% end if %>
		</tr>
	</table>

		<div class="div.contained_table">
      <table class="table table-custom table-responsive  table-condensed table-bordered">
			<% if (hscore > ascore) or (htToWon = true) or (htStlWon = true) or (ht3Won = true) then%>
			<thead>
			<th class="big" colspan="10"><i class="fas fa-check fa-lg" style="color: #468847;"></i>&nbsp;<span class="greenTrade"><%=objRBox.Fields("hometeam").Value%></span>&nbsp;<span style="color:darkorange;"><small>[<%=objRBox.Fields("hometeamshort").Value%>]</small></span></th>			
			</thead>	
			<%else%>
			<thead>
			<th style="color:#a5a6a7;" colspan="10"><%=objRBox.Fields("hometeam").Value%>&nbsp;<small>[<%=objRBox.Fields("hometeamshort").Value%>]</small></th>				
			</thead>	
			<%end if%>	
        <thead>
					<tr style="background-color:white;font-weight:bold;color:black;">
						<td class="big" style="vertical-align:middle;text-align:center;width:22%;"><span style="color:black;">PLAYER</span></td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">B</span>LK</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">A</span>ST</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">R</span>EB</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:10%"><span style="color:black;">P</span>TS</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">S</span>TL</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">3</span>PT</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">T</span>O</td>	
						<th class="big" width:20%;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i></th>			
					</tr>
        </thead>
				<tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">		
					<% if hp1Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=hp1PID %>"><%=left(hp1Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("hp1First").Value&"' AND Last = '"&objRBox.Fields("hp1Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGHP1 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
										
						'Response.Write "Season Barps  = "&seasonBarpAVGHP1&"<br>"
						'Response.Write "Compare Barps High = "&compareBarpsHigh&"<br>"
						'Response.Write "Compare Barps Low  = "&compareBarpsLow&"<br>"
						'Response.Write "FIRST         = "&objRBox.Fields("hp1First").Value&"<br>"
						'Response.Write "LAST          = "&objRBox.Fields("hp1Last").Value&"<br>"
						
						objRSBarps.Close					
						%>
					<%end if%>
						
						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp1BLK").Value) >=10 then %>
							<td class="big" style="font-size:11px;background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp1BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp1ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp1ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp1RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp1RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp1PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp1PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1PTS").Value)%></td>
						<%end if%>

						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp1STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp1STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1STL").Value)%></td>
						<%end if%>						

						<% if tbonushp1 >= 3 and CInt(objRBox.Fields("hp13S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;border-width: thin;"><%=(objRBox.Fields("hp13S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp13S").Value)%></td>
						<%end if%>	
						<td class="big"><%=(objRBox.Fields("hp1TOs").Value)%></td>

						<%if CDbl(objRBox.Fields("hp1Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("hp1Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("hp1Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("hp1Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("hp1Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("hp1Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp1Barps").Value)%></td>
						<%end if%>
						
					</tr>

				 <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if hp2Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=hp2PID %>"><%=left(hp2Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("hp2First").Value&"' AND Last = '"&objRBox.Fields("hp2Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp2 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
										
						'Response.Write "Season Barps  = "&seasonBarpAVGhp2&"<br>"
						'Response.Write "Compare Barps High = "&compareBarpsHigh&"<br>"
						'Response.Write "Compare Barps Low  = "&compareBarpsLow&"<br>"
						'Response.Write "FIRST         = "&objRBox.Fields("hp2First").Value&"<br>"
						'Response.Write "LAST          = "&objRBox.Fields("hp2Last").Value&"<br>"
						
						objRSBarps.Close					
						%>
					<%end if%>
						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp2BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp2BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp2ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp2ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp2RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp2RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp2PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp2PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2PTS").Value)%></td>
						<%end if%>

						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp2STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp2STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2STL").Value)%></td>
						<%end if%>						

						<% if tbonushp2 >= 3 and CInt(objRBox.Fields("hp23S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp23S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp23S").Value)%></td>
						<%end if%>	
						<td class="big"><%=(objRBox.Fields("hp2TOs").Value)%></td>
						
						<%if CDbl(objRBox.Fields("hp2Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("hp2Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("hp2Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("hp2Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("hp2Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("hp2Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp2Barps").Value)%></td>
						<%end if%>
					</tr>
					<tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if hp3Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=hp3PID %>"><%=left(hp3Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("hp3First").Value&"' AND Last = '"&objRBox.Fields("hp3Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp3 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if%>
						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp3BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp3BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp3ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp3ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp3RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp3RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp3PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp3PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3PTS").Value)%></td>
						<%end if%>

						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp3STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp3STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3STL").Value)%></td>
						<%end if%>						

						<% if tbonushp3 >= 3 and CInt(objRBox.Fields("hp33S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp33S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp33S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("hp3TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("hp3Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("hp3Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("hp3Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("hp3Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("hp3Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("hp3Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp3Barps").Value)%></td>
						<%end if%>
					</tr>
				  <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if hp4Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=hp4PID %>"><%=left(hp4Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("hp4First").Value&"' AND Last = '"&objRBox.Fields("hp4Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp4 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if%>
						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp4BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp4BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp4ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp4ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp4RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp4RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp4PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp4PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4PTS").Value)%></td>
						<%end if%>

						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp4STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp4STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4STL").Value)%></td>
						<%end if%>						

						<% if tbonushp4 >= 3 and CInt(objRBox.Fields("hp43S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp43S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp43S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("hp4TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("hp4Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("hp4Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("hp4Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("hp4Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("hp4Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("hp4Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp4Barps").Value)%></td>
						<%end if%>
					</tr>
				  <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">		
					<% if hp5Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=hp5PID %>"><%=left(hp5Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("hp5First").Value&"' AND Last = '"&objRBox.Fields("hp5Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if%>
						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp5BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp5BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp5ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp5ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp5RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp5RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5RBS").Value)%></td>
						<%end if%>
						
						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp5PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp5PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5PTS").Value)%></td>
						<%end if%>

						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp5STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp5STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5STL").Value)%></td>
						<%end if%>						

						<% if tbonushp5 >= 3 and CInt(objRBox.Fields("hp53S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("hp53S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp53S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("hp5TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("hp5Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("hp5Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("hp5Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("hp5Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("hp5Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("hp5Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("hp5Barps").Value)%></td>
						<%end if%>
					</tr>
					<% if (hscore > ascore) or (htToWon = true) or (htStlWon = true) or (ht3Won = true) then%>
						<tr style="font-size:11px;text-align:center;vertical-align:middle;font-weight:700;background-color:#468847;color:white">	
					<%else%>
						<tr style="font-size:11px;text-align:center;vertical-align:middle;font-weight:700;background-color:white">	
					<%end if%>
					<% if homePen > 0 then %>
          <td class="big" style="vertical-align:middle;" width="20%">PENALTY</td>
					<%else%>
          <td class="big" style="vertical-align:middle;" width="20%">TEAM</td>
					<%end if%>		
					<%if htBlkWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("htotBLK").Value)%></span></td>	
					<%else%>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotBLK").Value)%></td>
					<%end if%>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotASS").Value)%></td>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotRBS").Value)%></td>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotPTS").Value)%></td>
					<%if htStlWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("htotSTL").Value)%></span></td>
					<%else%>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotSTL").Value)%></td>
					<%end if %>
					<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htot3S").Value)%></td>
					<% if htToWon = true then %>
          <td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("htotTOs").Value)%></span></td>
					<%else %>
          <td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("htotTOs").Value)%></td>
					<%end if %>
						<% if (hscore > ascore) or (htToWon = true) or (htStlWon = true) or (ht3Won = true) then%>
					<td class="big" style="vertical-align:middle;text-align:center;background-color:#468847;;color:white;"><% =(objRBox.Fields("htotBarps").Value)%></td>
					<%else%>
						<td class="big"  style="vertical-align:middle;text-align:center;background-color:white"><% =(objRBox.Fields("htotBarps").Value)%></td>
				
					<%end if%>
        </tr>          
      </table>
	  </div>
		<div class="div.contained_table">
			<table class="table table-custom table-responsive  table-condensed table-bordered">
		<% if (ascore > hscore) or (atToWon = true) or (atStlWon = true) or (at3Won = true) then %>
			<thead>
				<th class="big" colspan="10"><i class="fas fa-check fa-lg" style="color: #468847;"></i>&nbsp;<span class="greenTrade"><%=objRBox.Fields("awayteam").Value%></span>&nbsp;<span style="color:darkorange;"><small>[<%=objRBox.Fields("awayteamshort").Value%>]</small></span></th>		
			</thead>	
			<%else%>
			<thead>
				<th style="color:#a5a6a7;" colspan="10"><%=objRBox.Fields("awayteam").Value%>&nbsp;<small>[<%=objRBox.Fields("awayteamshort").Value%>]</small></th>		
			</thead>
			<%end if%>
				<thead>
					<tr style="background-color:white;font-weight:bold;color:black;">
						<td class="big" style="vertical-align:middle;text-align:center;width:22%;"><span style="color:black;">PLAYER</span></td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">B</span>LK</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">A</span>ST</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">R</span>EB</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:10%"><span style="color:black;">P</span>TS</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">S</span>TL</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">3</span>PT</td>
						<td class="big" style="vertical-align:middle;text-align:center;width:8%"><span style="color:black;">T</span>O</td>	
						<th class="big" width:20%;"><span style="color:black;"><i class="fas fa-basketball-hoop"></i></th>			
					</tr>
				</thead>
				 <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<%if ap1Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else %>
						<td class="big" style="font-size:11px;color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=ap1PID %>"><%=left(ap1Last,11)%></a>						
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("ap1First").Value&"' AND Last = '"&objRBox.Fields("ap1Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
						
					<%end if%>
						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap1BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap1BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap1ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap1ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap1RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap1RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap1PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap1PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1PTS").Value)%></td>
						<%end if%>

						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap1STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap1STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1STL").Value)%></td>
						<%end if%>						

						<% if tbonusap1 >= 3 and CInt(objRBox.Fields("ap13S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap13S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap13S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("ap1TOs").Value)%></td>
						
						<%if CDbl(objRBox.Fields("ap1Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("ap1Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("ap1Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("ap1Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("ap1Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("ap1Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap1Barps").Value)%></td>
						<%end if%>
					
					</tr>
				 <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if ap2Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=ap2PID %>"><%=left(ap2Last,11)%></a>
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("ap2First").Value&"' AND Last = '"&objRBox.Fields("ap2Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10						
						objRSBarps.Close					
						%>
					<%end if %>
						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap2BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap2BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap2ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap2ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap2RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap2RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap2PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap2PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2PTS").Value)%></td>
						<%end if%>

						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap2STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap2STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2STL").Value)%></td>
						<%end if%>						

						<% if tbonusAP2 >= 3 and CInt(objRBox.Fields("ap23S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap23S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap23S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("ap2TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("ap2Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("ap2Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("ap2Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("ap2Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("ap2Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("ap2Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap2Barps").Value)%></td>
						<%end if%>
					</tr>
				  <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if ap3Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=ap3PID %>"><%=left(ap3Last,11)%></a>
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("ap3First").Value&"' AND Last = '"&objRBox.Fields("ap3Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if %>
						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap3BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap3BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap3ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap3ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap3RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap3RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap3PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap3PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3PTS").Value)%></td>
						<%end if%>

						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap3STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap3STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3STL").Value)%></td>
						<%end if%>						

						<% if tbonusap3 >= 3 and CInt(objRBox.Fields("ap33S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap33S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap33S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("ap3TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("ap3Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("ap3Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("ap3Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("ap3Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("ap3Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("ap3Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap3Barps").Value)%></td>
						<%end if%>
					</tr>
				 <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if ap4Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=ap4PID %>"><%=left(ap4Last,11)%></a>
						</td>
												<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("ap4First").Value&"' AND Last = '"&objRBox.Fields("ap4Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if %>
						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap4BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap4BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap4ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap4ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap4RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap4RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap4PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap4PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4PTS").Value)%></td>
						<%end if%>

						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap4STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap4STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4STL").Value)%></td>
						<%end if%>						

						<% if tbonusap4 >= 3 and CInt(objRBox.Fields("ap43S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap43S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap43S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("ap4TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("ap4Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("ap4Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("ap4Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><img class="img-responsive" style="max-width:100%;height:auto;margin:0px auto;display:block;border: #111;border-width: thin;border-style: solid;" src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcSEDYljPZt4ByFLb5hKtiYGlMntgcaiPNiWujzmjPTdj49h-9XkYQ">&nbsp;<%=(objRBox.Fields("ap4Barps").Value)%></td>
						<%elseif CDbl(objRBox.Fields("ap4Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("ap4Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap4Barps").Value)%></td>
						<%end if%>
					</tr>
				  <tr style="font-size:11px;font-weight:normal;background-color:black;color:white;">	
					<% if ap5Last = false then %>
						<td class="opp big" style="text-align:left">
							<i class="fas fa-user-alt-slash big"></i>
						</td>
					<%else%>
						<td class="big" style="color:#01579B;font-weight:bold;text-align:left;">
							<a class="white" href="playerprofile.asp?pid=<%=ap5PID %>"><%=left(ap5Last,11)%></a>
						</td>
						<%     
						Set objrsBarps = Server.CreateObject("ADODB.RecordSet")
						compareBarpsHigh      = 0
						compareBarpsLow       = 0
						
						objRSBarps.Open	"SELECT * FROM tbl_barps " & _
	                  "WHERE First = '"&objRBox.Fields("ap5First").Value&"' AND Last = '"&objRBox.Fields("ap5Last").Value&"' ", objConn,3,3,1	
										seasonBarpAVGhp5 = objRSBarps.Fields("barps").Value									
										compareBarpsHigh      = objRSBarps.Fields("barps").Value + 10
										compareBarpsLow       = objRSBarps.Fields("barps").Value - 10
						objRSBarps.Close					
						%>
					<%end if %>
						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap5BLK").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap5BLK").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5BLK").Value)%></td>
						<%end if%>						
						
						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap5ASS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap5ASS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5ASS").Value)%></td>
						<%end if%>	
	
						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap5RBS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap5RBS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5RBS").Value)%></td>
						<%end if%>
						
						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap5PTS").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap5PTS").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5PTS").Value)%></td>
						<%end if%>

						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap5STL").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap5STL").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5STL").Value)%></td>
						<%end if%>						

						<% if tbonusap5 >= 3 and CInt(objRBox.Fields("ap53S").Value) >=10 then %>
							<td class="big" style="background-color:darkorange;color: white;font-weight:bold;border: black;border-style: solid;    border-width: thin;"><%=(objRBox.Fields("ap53S").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap53S").Value)%></td>
						<%end if%>	
						
						<td class="big"><%=(objRBox.Fields("ap5TOs").Value)%></td>
						<%if CDbl(objRBox.Fields("ap5Barps").Value) = 0 then %>
							<td class="big"><%=(objRBox.Fields("ap5Barps").Value)%></td>								
						<%elseif CDbl(objRBox.Fields("ap5Barps").Value) > CDbl(compareBarpsHigh) then%>
							<td class="big"><span><i class="fas fa-fire fire"></i>&nbsp;<%=(objRBox.Fields("ap5Barps").Value)%></span></td>
						<%elseif CDbl(objRBox.Fields("ap5Barps").Value) < CDbl(compareBarpsLow) then%>
							<td class="big"><i class="far fa-snowflake" style="color:cornflowerblue;font-weight:bold;"></i>&nbsp;<%=(objRBox.Fields("ap5Barps").Value)%></td>
						<%else%>
							<td class="big"><%=(objRBox.Fields("ap5Barps").Value)%></td>
						<%end if%>
				</tr>
				<% if (ascore > hscore) or (atToWon = true) or (atStlWon = true) or (at3Won = true) then %>
			  <tr style="font-size:11px;text-align:center;vertical-align:middle;font-weight:700;background-color:#468847;color:white;">		
				<%else%>
				 <tr style="font-size:11px;text-align:center;vertical-align:middle;font-weight:700;background-color:white;">		
				<%end if%>	
						<% if awayPen > 0 then %>
						<td class="big" style="vertical-align:middle;" width="20%">PENALTY</td>
						<%else%>
						<td class="big" style="vertical-align:middle;" width="20%">TEAM</td>
						<%end if%>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotBLK").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotASS").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotRBS").Value)%></td>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotPTS").Value)%></td>							
							<%if atStlWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("atotSTL").Value)%></span></td>
							<%else%>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotSTL").Value)%></td>
							<%end if%>
							<%if atBlkWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("atot3S").Value)%></span></td>		
							<%else%>	
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atot3S").Value)%></td>	
							<%end if%>
							<% if atToWon = true then %>
							<td class="big" style="vertical-align:middle;"><span class="badgeTie"><%=(objRBox.Fields("atotTOs").Value)%></span></td>
							<%else%>
							<td class="big" style="vertical-align:middle;"><%=(objRBox.Fields("atotTOs").Value)%></td>
							<%end if%>
							<td class="big" ><% =(objRBox.Fields("atotBarps").Value)%></td>
				</tr>
			</table>
</br>			
								</div>

