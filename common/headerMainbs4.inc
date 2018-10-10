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
<nav class="navbar navbar-expand-lg navbar-light bg-light">
	<a href="dashboard.asp" class="navbar-brand navbar-left"><i class="fa fa-home" aria-hidden="true"></i>&nbsp;<strong><%=objRSteamsSession("TeamName").value%></strong></a>
  <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
    <span class="navbar-toggler-icon"></span>
  </button>
  <div class="collapse navbar-collapse" id="navbarSupportedContent">
    <ul class="navbar-nav mr-auto">
      <li class="nav-item dropdown">
        <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
          Players
        </a>
				<div class="dropdown-menu" aria-labelledby="navbarDropdown">	
					<%if (wTradeDeadLine + 1 + 1/24) < now() then %>  
					<% else %>
						<% if myTradeInd = false then%>
						<%else%>
							<a class="dropdown-item" href="tradeanalyzer.asp" target="_self"><i class="fas fa-exchange-alt"></i>&nbsp;Submit Trade Offers</a>
						<%end if%>			
					<%end if%>				
					
					<% if w_seasonOver = true and 1=2 then %>
						<a class="dropdown-item" href="" target="_self"><i class="fas fa-frown"></i>&nbsp;Your Season is Over!</a>
					<%else%>
						<a class="dropdown-item" href="transelect.asp" target="_self"><i class="fa fa-user" aria-hidden="true"></i>&nbsp;Acquisition</a>
						<a class="dropdown-item" href="playerSearch.asp" target="_self"><i class="fa fa-search" aria-hidden="true"></i>&nbsp;Search</a>
						<a class="dropdown-item" href="barpListing.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Comparison</a>
						<a class="dropdown-item" href="auctionrecap.asp" target="_self"><i class="fas fa-usd-circle"></i>&nbsp;Auction Recap</a>
					<%end if%>  
        </div>
      </li>
      <li class="nav-item dropdown">
        <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
          STATS & INFO
        </a>
        <div class="dropdown-menu" aria-labelledby="navbarDropdown">
					<a class="dropdown-item" href="results.asp"><i class="fa fa-calculator" aria-hidden="true"></i>&nbsp;Results</a>
					<a class="dropdown-item" href="allstandings.asp"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Standings</a>
					<a class="dropdown-item" href="finreports.asp" target="_self"><i class="fa fa-usd" aria-hidden="true"></i>&nbsp;Transactions and Financials</a>
					<a class="dropdown-item" href="allskeds.asp"><i class="fal fa-calendar-alt"></i>&nbsp;Schedules</a>
					<a class="dropdown-item" href="gameGridAll.asp"><i class="fal fa-calendar-alt"></i>&nbsp;Game Grid</a>
					<a class="dropdown-item" href="teamanalysis.asp"><i class="fa fa-bolt" aria-hidden="true"></i>&nbsp;Team Analysis</a>
					<a class="dropdown-item" href="emailtheleague.asp"><i class="fa fa-pencil-square-o" aria-hidden="true"></i>&nbsp;Email the League</a>
					<a class="dropdown-item" href="contacts.asp"><i class="fa fa-users"></i>&nbsp;Contacts</a>
					<a class="dropdown-item" href="championship_alley.asp"><i class="fa fa-trophy" aria-hidden="true"></i>&nbsp;Champions Row</a>
          <div class="dropdown-divider"></div>
          <a class="dropdown-item" href="rules.asp"><i class="fa fa-question-circle" aria-hidden="true"></i>&nbsp;HELP!</a>
					<a class="dropdown-item" href="login.asp"><i class="fa fa-question-circle" aria-hidden="true"></i>&nbsp;Logout!</a>
        </div>
      </li>
			<% if ownerid = 10 or ownerid = 8 then %>
      <li class="nav-item dropdown">
        <a class="nav-link dropdown-toggle" href="#" id="navbarDropdown" role="button" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
          Commish Tools
        </a>
        <div class="dropdown-menu" aria-labelledby="navbarDropdown">
					<a class="dropdown-item href="transferLast5.asp"><i class="fas fa-exchange-alt"></i>&nbsp;Xfer Last Five</a>
					<a class="dropdown-item href="BuildBox.asp"><i class="fa fa-calculator"></i>&nbsp;Generate Results</a>	
					<a class="dropdown-item href="playerAdd.asp"><i class="fa fa-user"></i>&nbsp;Add Player</a>	
					<a class="dropdown-item href="commishMatchups.asp"><i class="fa fa-calculator"></i>&nbsp;Matchups</a>
					<a class="dropdown-item href="playerMaint.asp"><i class="fa fa-users"></i>&nbsp;Player Maintenance</a>
					<a class="dropdown-item href="ownerMaint.asp"><i class="fa fa-users"></i>&nbsp;Owner Maintenance</a>
					<a class="dropdown-item href="lineupHistory.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Lineup History</a>
					<a class="dropdown-item href="tieBreakerSked.asp"><i class="fa fa-sort-numeric-asc" aria-hidden="true"></i>&nbsp;Update Rankings</a>
					<a class="dropdown-item href="maintainSchedule.asp"><i class="fa fa-list" aria-hidden="true"></i>&nbsp;Configure Position Rnd</a>
					<a class="dropdown-item href="parmMaint.asp"><i class="fa fa-cog" aria-hidden="true"></i>&nbsp;Configure Parameters</a>
					<a class="dropdown-item href="maintainTeamSkeds.asp"><i class="fa fa-cog" aria-hidden="true"></i>&nbsp;Update Game Times</a>
					<a class="dropdown-item href="time.asp"><i class="fa fa-clock-o" aria-hidden="true"></i>&nbsp;Update Timed Events</a>
					<a class="dropdown-item href="loadrosters.asp"><i class="fa fa-users" aria-hidden="true"></i>&nbsp;Load Rosters</a>
        </div>
			</li>
			<%end if%>   
		</ul>
  </div>
</nav>