<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""
	
  dim objConn,objRSWork
	
	Set objConn        = Server.CreateObject("ADODB.Connection")

	objConn.Open Application("lineupstest_ConnectionString")
	objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
								"Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
				  
	%>
	<!--#include virtual="Common/session.inc"-->

<!DOCTYPE HTML>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
gray {
	font-size:12px;
}
red {
	color: red;
}
black {
	color: black;
}
white {
	color: white;
}
green {
	color: #468847;
		font-size:12px;
}
.alert-warning {
    color: #8a6d3b;
    background-color: #fcf8e3;
    border-color: #01579B;
}
.panel-override {
    color: black;
    background-color:WHITE;
    border-color: black;
    border-radius: 3px!important;
}
panel-title {
    color: yellowgreen;
    text-transform: none;
    font-size: 20px !important;
    font-weight: 700;
}
red {
	color: red;
	font-weight:700;;
}
</style>
</head>
<body>
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong>Rules and Settings</strong>
			</div>
		</div>
	</div>
</div>
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="panel panel-override">
				<table class="table table-custom-black table-responsive table-bordered table-condensed">
				<tr>
					<th colspan="2">ENTRY FEES</th>
				</tr>
				<tr><td colspan="2">Entry Fee: $185.60 per owner for Team </td></tr>
				<tr><td colspan="2">Site Hosting Fee: $14.40 per team. Due October 24th.</td></tr>
				<tr>
					<th colspan="2">AUCTION</th>
				</tr>
				<tr><td colspan="2"><a class="blue big" href="http://fantasy.espn.com/basketball/league?leagueId=66401">Date:&nbsp;Monday, October 22nd at 8:00 PM CST</a></td></tr>
				<tr><td colspan="2">Auction Budget:&nbsp;$200 to Spend on 14 Players</td></tr>
				<tr>
					<th colspan="2">ROSTERS</th>
				</tr>	
				<tr><td colspan="2">14 Players max with no limit on player positions on a roster. Player positions will be designated prior to the Auction based on <span style="color: red;font-weight: bold">IGBL Position Guidelines.</span></td></tr>
				<tr><th colspan="2">OBJECT OF THE GAME:</th></tr>
				<tr>
					<td colspan="2">To outscore your opponent in head to head play. Your opponents will be predetermined for each game and posted on the IGBL schedule. The regular season consists of 7 intervals of 10 games each.
								After playing each team once in league round robin play, a position round game will be played to complete the interval.
					</td>
				</tr>
				<tr>
					<th colspan="2">POSITION ROUND</th>
				</tr>
				<tr>
					<td colspan="2" style="text-align:center;">The league position round will pit seeds </br></br>1 vs 2</br>  3 vs 4</br>  5 vs 6</br>  7 vs 8</br>  9 vs 10</br></br>based on current day standings.  
					</td>
				</tr>
				<tr>
					<th colspan="2">LINE-UPS</th>
				</tr>
				<tr>
					<td colspan="2">Player eligibility sets for each player Individually at Scheduled Game-time tip of their NBA team.</td>
				</tr>
				<tr>
					<td style="width:50%;">Center</td>
					<td style="width:50%;text-align:center;">1 </td>
				</tr>
				<tr>
					<td style="width:50%;">Forwards</td>
					<td style="width:50%;text-align:center;">2 </td>
				</tr>
				<tr>
					<td style="width:50%;">Guards</td>
					<td style="width:50%;text-align:center;">2</td>
				</tr>
				<tr>
					<td>Penalty</td>
					<td><red>Failure to submit a lineup results in an automatic loss and your opponent must beat your optimal lineup based on available players. If the penalized owner's lineup wins the game both teams are awarded a loss.</red></td>
				</tr>
				<tr>
					<th colspan="2">SCORING RULES</th>
				</tr>
					<tr>
						<td style="width:50%;">Blocked Shots</td>
						<td style="width:50%;text-align:center;">2</td></tr>
					<tr>
						<td style="width:50%;">Assists</td>
						<td style="width:50%;text-align:center;">1.5</td>
					</tr>
					<tr>
						<td style="width:50%;">Rebounds</td>
						<td style="width:50%;text-align:center;">1.25</td>
					</tr>
					<tr>
						<td style="width:50%;">Points Scored</td>
						<td style="width:50%;text-align:center;">1</td>
					</tr>
					<tr>
						<td style="width:50%;">Steals </td>
						<td style="width:50%;text-align:center;">2</td>
					</tr>
					<tr>
						<td style="width:50%;">3 Pointers Made </td>
						<td style="width:50%;text-align:center;">1</td>
					</tr>
					<tr>
						<td style="width:50%;">Triple Double</td>
						<td style="width:50%;text-align:center;">3</td>
					</tr>
					<tr>
						<td style="width:50%;">Turnovers </td>
						<td style="width:50%;text-align:center;">-1</td>
					</tr>
					<tr>
						<th colspan="2">GAME TIE BREAKER RULES</th>
					</tr>
					<tr>
						<td colspan="2">
								<ol>
									<li>Team with the least turnovers</li>
									<li>Team with the most steals</li>
									<li>Team with the most blocks</li>
									<li>Team with the most 3-pointers</li>
									<li>Team with the most rebounds</li>
									<li>Team with the most assists</li>
									<li>Team with the most points</li> 
									<li>Home Team</li>
									<li><span class="badgeDown">Category Used to Break the Tie</span></li>
								</ol>
						</td>
					</tr>
					<tr>
						<th colspan="2">STANDINGS TIE-BREAKER RULES</th>
					</tr>
					<tr>
						<td colspan="2">
							<ol>
								<li>The Rules Below Enforced on Position Round!</li>
								<li>If multiple teams share identical records, it's determined if an owner has a H2H advantage over all tied teams. If so that team will be placed highest in the standings</li>
								<li>If Step 2 does not render a decision teams are ranked based on P/PG from (Hi/Lo)</li>
								<li>Repeat Step 2 if necessary until all tied teams are ranked</li>
							</ol>
						</td>
					</tr>
					<tr> 
						<th COLSPAN="2">WAIVER PROCESS</th>
					</tr>
						 <tr style="background-color:#FFFFFF">	
						<td>
							<ol>
								<li>Each team starts the season with $250 waiver budget </li>
								<li>Waivers will run as they do today with players being awarded to the highest bidder</li>
								<li>Upon a successful waiver claim the owner will be moved to the bottom of waiver priority</li>
								<li>If multiple owners submit the same bid and it is the highest bid, the owner with the highest waiver priority will be awarded the player</li>
								<li>All other players will be awarded in the same manner until all waiver requests  have been exhausted </li>
								<li>Min Waiver Bid $1 - Max Waiver Bid $Waiver Bal</li>
								<li>Once your wavier bank roll has been exhausted, you won't be able to acquire players via waivers, only via the Free Agency period where you can Rent/Acquire Free Agents (BAU)</li>
								<li>Waivers are processed starting with the highest Bid submitted</li>
							</ol>
						</td>
						<td>
<strong>FREE AGENCY AUCTION BIDDING</strong>
Bring the auction experience to your free agent list. A free-agent acquisition budget ($250) is used in blind auctions open to every team within the league.

In an auction bidding environment, players who are not selected in the league draft, players who are dropped and players added to the player pool during the season become eligible to be claimed via waivers. However, instead of placing a traditional waiver claim, an owner places a bid they feel is appropriate based on a player's value. During the bidding period, team owners may bid at any time before the deadline. In a bidding environment, the only way players can be picked up is through the bidding process, and there are no first-come, first-served pick-ups. The bidding process is an open process and there is no sequence for the bids. Team owners may bid at any time during the waiver period, before it expires.

In the event of bids of equal amounts for a player by two or more teams (a tie), the team with the highest waiver priority will receive the player. During the bidding period, no team can see any other team's bids or bid amounts. All teams can view bidding results after bids are processed via the waiver reports. Additionally, all teams can view other teams' budget balances on each team clubhouse and waiver order pages.  						
						</td>
					</tr>				 
			
					<%
				Set objRSWork          = Server.CreateObject("ADODB.RecordSet") 
					
				objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'PICKUP' ",objConn
				wPickUp = objRSWork.Fields("param_amount").Value
				objRSWork.Close

				objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE' ",objConn
				wTrade = objRSWork.Fields("param_amount").Value
				objRSWork.Close
				
				objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT' ",objConn
				wRental = objRSWork.Fields("param_amount").Value
				objRSWork.Close
				
				objRSWork.Open "SELECT * FROM tblParameterCtl where param_name = 'RENT_PENALTY' ",objConn
				wRentPenalty = objRSWork.Fields("param_amount").Value
				objRSWork.Close
				%>
					<tr> 
						<th style="width:50%;">TRANSACTIONS</th>
						<th style="width:50%;text-align:center">Cost</th>
					</tr>
					 <tr style="background-color:#FFFFFF">	
						<td>Free Agent Acquisition</td>
						<td style="text-align:center">$<%= wPickUp%>.00</td>
					</tr>
					 <tr style="background-color:#FFFFFF">	
						<td>Rental Acquisition</td>
						<td style="text-align:center">$<%= wRental%>.00</td>
					</tr>
					<tr style="background-color:#FFFFFF">	
						<td>Rental Acquisition Penalty</td>
						<td style="text-align:center">$<%= wRentPenalty%>.00</td>
					</tr>
					 <tr style="background-color:#FFFFFF">	
						<td>Trade Acquisition</td>
						<td style="text-align:center">$<%= wTrade%>.00</td>
					</tr>
					 <tr style="background-color:#FFFFFF">
						<td>Waiver Acquisition</td>
						<td style="text-align:center">$<%= wPickUp%>.00</td>
					</tr>
					<tr> 
						<th colspan ="2" style="width:50%;">TRADE ANALYSIS SUMMARY</th>
					</tr>
					 <tr style="background-color:#FFFFFF">	
						<td colspan="2"><strong>Positive</strong> <greenIcon><i class="fa fa-thumbs-up" aria-hidden="true"></i></greenIcon></br>
						Number of Lineups with Positive Impact
						</td>
					</tr>
					 <tr style="background-color:#FFFFFF">	
						<td colspan="2"><strong>Negative</strong> <redIcon><i class="fa fa-thumbs-down" aria-hidden="true"></i></redIcon></br>
						Number of Lineups with Negative Impact
						</td>
					</tr>
					<tr style="background-color:#FFFFFF">	
						<td colspan="2"><strong>Neutral [+/- 5]</strong> <evenIcon><i class="fa fa-balance-scale" aria-hidden="true"></i></evenIcon></br>
						Barp Threshold Used to Determine Positive/Negative Impact on Your Lineup. If threshold = 5, a Barp difference <=5 will count as <strong>No Lineup Impact</strong> and the Neutral Count will be incremented +1 when this condition exists.</td>
					</tr>
					<tr>
						<th colspan="2">BARP BADGES | Background Colors</th>
					</tr>
					<tr style="background-color:#01579b;">
						<td colspan="2" style="font-weight:bold;"><span class="badgeBlue">Season Barp Average</span></td>
					</tr>
					<tr style="background-color:black;">
						<td colspan="2" style="font-weight:bold;"><span class="badgeHomeTeam">Season Barp Average</span></td>
					</tr>
					<tr style="background-color:#468847;">
						<td colspan="2" style="font-weight:bold;"><span class="badgeUp">Barp Average is&nbsp;<i class="far fa-long-arrow-up fa-lg"></i>&nbsp;Over Last 5 Games</span></td>
					</tr>
					<tr style="background-color:#9a1400;">
						<td colspan="2" style="font-weight:bold;"><span class="badgeDown">Barp Average is&nbsp;<i class="far fa-long-arrow-down fa-lg"></i>&nbsp;DOWN Over Last 5 Games</span></td>
					</tr>
					<tr style="background-color:gold;">
						<td colspan="2" style="font-weight:bold;"><span class="badgeEven">Barp Average is&nbsp;<i class="fal fa-arrows-h fa-lg"></i>&nbsp;STABLE Over Last 5 Games</span></td>
					</tr>
						<tr>
						<th colspan="2">ICONS</th>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><redText><i class="far fa-file-invoice-dollar fa-lg"></i></redText>&nbsp;View Transaction Report(s)</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><redText><i class="far fa-inbox-in fa-lg"></i></redText>&nbsp;Trade Offers Received</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><redText><i class="far fa-inbox-out fa-lg"></i></redText>&nbsp;Trade Offers Awaiting Owners Response</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><redText><i class="fa fa-scissors fa-lg"></i></redText>&nbsp;Pending Waivers</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><redText><i class="fa fa-balance-scale fa-lg"></i></redText>&nbsp;Trades Being Analyzed</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><strong><i class="fa fa-registered blue" aria-hidden="true"></i></strong>&nbsp;Player Rented</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><strong><i class="fas fa-briefcase-medical red"></i>&nbsp;Player On IR</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><strong><i class="far fa-exchange"></i></strong>&nbsp;Player Trade Pending</td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"><strong><i class="fas fa-user-clock auctionText"></i></strong>&nbsp;Player Waiver Pending</td>
					</tr>
					<tr>
						<th colspan="2">ALERTS</th>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">RECEIVE ALERTS VIA EMAIL:<br><span style="font-weight:600;color:#303F9F;">Alerts delivered via email</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">RECEIVE ALERTS VIA TEXT:<br><span style="font-weight:600;color:#303F9F;">Alerts delivered via text</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">FREE AGENT SIGNINGS:<br><span style="font-weight:600;color:#303F9F;">Alerted when free-agent is signed</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">PLAYER RENTALS:<br><span style="font-weight:600;color:#303F9F;">Alerted when player is rented</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">TRADES:<br><span style="font-weight:600;color:#303F9F;">Alerted of (Accepts, Offers, Rejects & Withdrawals) </span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">WAIVERS RUN:<br><span style="font-weight:600;color:#303F9F;">Alerted when waivers have processed</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;"> STAGGER WINDOW (O/C):<br><span style="font-weight:600;color:#303F9F;">Alerted when stagger window open/closes</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">BOX SCORES GENERATED:<br><span style="font-weight:600;color:#303F9F;">Alerted when scores/standings have been updated</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">ON THE BLOCK POSTS:<br><span style="font-weight:600;color:#303F9F;">Alerted of OTB posts</span></td>
					</tr>
					<tr>
						<td colspan="2" style="font-weight:bold;">EMAIL THE LEAGUE POSTS:<br><span style="font-weight:600;color:#303F9F;">Alerted when email sent to the league</span></td>
					</tr>
					<tr> 
						<th colspan="2">PLAYOFFS | WINNING SHARES</th>
					</tr>
					<tr style="background-color:#FFFFFF">	
						<td>Regular Season Winner</td>
						<td style="text-align:center">$200.00</td>
					</tr>
					<tr style="background-color:#FFFFFF">	
						<td>Cycle Winner(s) [7]</td>
						<td style="text-align:center">$25 </td>
					</tr>
					<tr style="background-color:#FFFFFF">	
						<td>1st Round Byes </td>
						<td style="text-align:center">1st & 2nd Place Finishers</td>
					</tr>
					 <tr style="background-color:#FFFFFF">
						<td>Quarter-Finals</td>
						<td style="text-align:center">3rd Place vs 6th Place <br>4th Place vs 5th Place</td>
					</tr>
					 <tr style="background-color:#FFFFFF">
						<td>Semi-Finals</td>
						<td style="text-align:center">1st Place vs Lowest Seed <br>2nd Place vs Remaining Team</td>
					</tr>
					 <tr style="background-color:#FFFFFF">
						<td>Finals</td>
						<td style="text-align:center">Winners of Semi-Finals Series</td>
					</tr>
					 <tr style="background-color:#FFFFFF">
						<td>Consolation</td>
						<td style="text-align:center">Losers of Semi-Finals Series</td>
					</tr>
					<tr style="background-color:#FFFFFF">
						<td>Payouts</td>
						<td style="text-align:center">1st Place 60%<br>2nd Place 30%<br>3rd Place 10%</td>
					</tr>
					</table>
			</div>
		</div>
	</div>
</div>

<script language="JavaScript" type="text/javascript">
$(document).ready(function(){
    $('[data-toggle="tooltip"]').tooltip(); 
});		
</script>
</body>
</html>
