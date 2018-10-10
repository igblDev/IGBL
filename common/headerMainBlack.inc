<%
Set objRSmainwrk    = Server.CreateObject("ADODB.RecordSet")
Set objParams       = Server.CreateObject("ADODB.RecordSet")
Set objRSMe         = Server.CreateObject("ADODB.RecordSet")

objRSmainwrk.Open "SELECT * FROM tblParameterCtl where param_name = 'PLAYOFF_START_DATE' ",objConn
wMainPlayoffStart = objRSmainwrk.Fields("param_date").Value
objRSmainwrk.Close

objParams.Open "SELECT * FROM tblParameterCtl where param_name = 'TRADE_DEADLINE' ",objConn
wTradeDeadLine = objParams.Fields("param_date").Value
objParams.Close

objRSMe.Open "SELECT * FROM tblOwners where ownerID = "&ownerid&" ", objConn,3,3,1 
myTradeInd = objRSMe.Fields("acceptTradeOffers").Value
objRSMe.Close
%>	  
<header class="navbar navbar-custom navbar-inverse navbar-fixed-top bs-docs-nav" role="banner">
	<div class="container">
		<div class="navbar-header">
			<button class="navbar-toggle" type="button" data-toggle="collapse" data-target=".bs-navbar-collapse"> 
				<span class="sr-only">Toggle navigation</span> 
				<span class="icon-bar"></span> 
				<span class="icon-bar"></span> 
				<span class="icon-bar"></span> 
			</button>
			<a href="myigbl.asp" class="navbar-brand navbar-left"><i class="fa fa-home" aria-hidden="true"></i>&nbsp;<strong><%=objRSteamsSession("TeamName")%></strong></a> 
		</div>
		<nav class="collapse navbar-collapse bs-navbar-collapse" role="navigation">
			<ul class="nav navbar-nav">
				<li class="dropdown"> <a href="#" class="dropdown-toggle" data-toggle="dropdown">Players <b class="caret"></b></a>
            <ul class="dropdown-menu" role="menu">
              <li><a href="transelect.asp" target="_self"><i class="fa fa-user" aria-hidden="true"></i>&nbsp;Player Acquisition</a></li>
							<% if (wTradeDeadLine + 1 + 1/24) < now() then %>  
							<% else %>
								<% if myTradeInd = false then%>
								<%else%>
									<li><a href="tradeanalyzer.asp" target="_self"><i class="fas fa-exchange-alt"></i>&nbsp;Submit Trade Offers</a></li>
								<%end if%>
							<li><a href="ontheblock.asp" target="_self"><i class="fa fa-shopping-cart" aria-hidden="true"></i>&nbsp;OTB</a></li>
							<%end if%>
							<li><a href="playerSearch.asp" target="_self"><i class="fa fa-search" aria-hidden="true"></i>&nbsp;Search</a></li>
							<li><a href="maintainRoster.asp" target="_self"><i class="fa fa-road" aria-hidden="true"></i>&nbsp;Waive Players</a></li>
              <!--<li><a href="waivers_priority.asp"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Waiver Priority</a></li>-->
            </ul>
				</li>
				<li class="dropdown"> <a href="#" class="dropdown-toggle" data-toggle="dropdown">Reports <b class="caret"></b></a>
					<ul class="dropdown-menu" role="menu">
						<li><a href="finreports.asp" target="_self"><i class="fa fa-usd" aria-hidden="true"></i>&nbsp;Transactions and Financials</a></li>
						<li><a href="rosters.asp"><i class="fa fa-users" aria-hidden="true"></i>&nbsp;Rosters</a></li>
						<li><a href="barpListing.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Player Comparison</a></li>
					</ul>
				</li>
				<li class="dropdown"> <a href="#" class="dropdown-toggle" data-toggle="dropdown">Stats & Info <b class="caret"></b></a>
					<ul class="dropdown-menu" role="menu">
						<li><a href="viewLineups.asp" target="_self"><i class="fas fa-basketball-ball"></i>&nbsp;Lineups <i class="fas fa-ticket-alt"></i> Matchups</a></li>              
						<li><a href="results.asp"><i class="fa fa-calculator" aria-hidden="true"></i>&nbsp;Results</a></li>
						<!--<li><a href="gamelog.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Game Logs</a></li>-->
						<li><a href="allstandings.asp"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Standings</a></li>
						<li><a href="allskeds.asp"><i class="fa fa-calendar" aria-hidden="true"></i>&nbsp;Schedules</a></li>
						<li><a href="gameGridAll.asp"><i class="fa fa-calendar" aria-hidden="true"></i>&nbsp;Game Grid</a></li>
						<li><a href="teamanalysis.asp"><i class="fa fa-bolt" aria-hidden="true"></i>&nbsp;Strenghts & Weaknesses</a></li>
						<li><a href="championship_alley.asp"><i class="fa fa-trophy" aria-hidden="true"></i>&nbsp;Champions Row</a></li>
						<li><a href="emailtheleague.asp"><i class="fa fa-pencil-square-o" aria-hidden="true"></i>&nbsp;Email the League</a></li>
						<li><a href="contacts.asp"><i class="fa fa-users"></i>&nbsp;Contacts</a></li>
						<li><a href="rules.asp"><i class="fa fa-question-circle" aria-hidden="true"></i>&nbsp;HELP!</a></li>
						</ul>
				</li>
				<% if ownerid = 10 or ownerid = 8 then %>
				<li class="dropdown"> <a href="#" class="dropdown-toggle" data-toggle="dropdown">Commish <b class="caret"></b></a>
					<ul class="dropdown-menu" role="menu">
						<li><a href="transferLast5.asp"><i class="fas fa-exchange-alt"></i>&nbsp;Xfer Last Five</a></li>
						<li><a href="BuildBox.asp"><i class="fa fa-calculator"></i>&nbsp;Generate Results</a></li>	
						<li><a href="playerAdd.asp"><i class="fa fa-user"></i>&nbsp;Add Player</a></li>	
						<li><a href="commishMatchups.asp"><i class="fa fa-calculator"></i>&nbsp;Matchups</a></li>
						<li><a href="playerMaint.asp"><i class="fa fa-users"></i>&nbsp;Player Maintenance</a></li>
						<!--<li><a href="barpMaint.asp"><i class="fa fa-user"></i>&nbsp;UPT Name On Last 5</a></li>-->
						<li><a href="lineupHistory.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Lineup History</a></li>
						<li><a href="tieBreakerSked.asp"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Update Rankings</a></li>
						<li><a href="maintainSchedule.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Configure Position Rnd</a></li>
						<li><a href="parmMaint.asp"><i class="fa fa-cog" aria-hidden="true"></i>&nbsp;Configure Parameters</a></li>
						<li><a href="maintainTeamSkeds.asp"><i class="fa fa-cog" aria-hidden="true"></i>&nbsp;Update Game Times</a></li>
						<li><a href="time.asp"><i class="fa fa-clock-o" aria-hidden="true"></i>&nbsp;Update Timed Events</a></li>
						<li><a href="loadrosters.asp"><i class="fa fa-users" aria-hidden="true"></i>&nbsp;Load Rosters</a></li>
						</ul>
				</li>
				<%end if %>
			</ul>			
		</nav>
	</div>
</header>