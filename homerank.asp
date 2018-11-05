<!-- #include file="adovbs.inc" -->
<!--#include virtual="../Common/IGBLStandard.inc"-->
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

 	Dim objRSgames,objRS,objConn, objRSwaivers
	Dim strSQL, iPlayerClaimed,objRSTxns, objRSOwners, objRejectWaivers, iPlayerWaived, iOwner, w_action
   		
	Set objConn  = Server.CreateObject("ADODB.Connection")
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	Set objRSgames = Server.CreateObject("ADODB.RecordSet")
	Set objRSwaivers = Server.CreateObject("ADODB.RecordSet")
	Set objRSTxns 		= Server.CreateObject("ADODB.RecordSet")
	Set objRejectWaivers	= Server.CreateObject("ADODB.RecordSet")   			
	Set objNextRun	= Server.CreateObject("ADODB.RecordSet")   			

	
	objConn.Open Application("lineupstest_ConnectionString")


    objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
	              "Persist Security Info=False"
	 
	'*************************************************
	'Run setwaiverpriority event if it hasn't been run today.
	'*************************************************
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'setwaiverpriority' and nextrun < now() ", objConn,3,3,1
	                  
	if  objRSwaivers.Recordcount > 0 then
		objRSTxns.Open	"SELECT o.OwnerID, o.waiverpriority, o.TeamName, s.won, s.ppg " & _
   		                    "FROM standings s, tblOwners o " & _
   		                    "WHERE s.ID = o.ownerid " & _
   		                    "ORDER by s.won, s.ppg, s.oppg desc ", objConn
		w_priority = 1
			
		While Not objRSTxns.EOF
		   	
			iOwner = objRSTxns.Fields("OwnerID").Value
				
			strSQL = "update tblOwners " & _
		    	     "SET waiverPriority = " & w_priority & " " & _
			         "WHERE OwnerID = " &iOwner& "  ; "
			objConn.Execute strSQL
				
			'Response.Write "Sql = " & strSQL  & ".<br>"

			w_priority = w_priority + 1
			objRSTxns.MoveNext
   		Wend
   		
		objRSTxns.Close
		
		strSQL = "update tbltimedEvents " & _
		        "SET lastrun = now(), nextrun = nextrun + 1 " & _
		        "WHERE event = 'setwaiverpriority' "
		objConn.Execute strSQL		
		
	end if
	 
	objRSwaivers.Close	
	 
	'********************************************************
	'Run pendingwaiversall event if it hasn't been run today.
	'********************************************************
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'pendingwaiversall' and nextrun < now() ", objConn,3,3,1
	                  
	if  objRSwaivers.Recordcount > 0 then
		objRSTxns.Open		"SELECT * FROM qryUpdatewaiver ", objConn,3,3,1
    	w_action = objRSTxns.Recordcount
	
		'Response.Write "Count = : " & w_action & "<br>"
    	while w_action > 0 
   				 			
	   		iPlayerClaimed= objRSTxns.Fields("PID_Claimed").Value
			iPlayerWaived = objRSTxns.Fields("PID_Waived").Value
			iOwner = objRSTxns.Fields("OwnerID").Value
			iPriority = objRSTxns.Fields("WaiverPriority").Value
			iActivePlayers = objRSTxns.Fields("ActivePlayerCnt").Value

			if iPlayerWaived = 0 AND iActivePlayers >= 15 then
				'***************************************************************
				'REJECT THIS TRANSACTION BECAUSE THE PLAYER LIMIT IS 15 PER TEAM
				'****************************************************************
				TransType = "Waiver pickup rejected (Roster full)"
				Cost = 0.00			
					
				strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_ 
				iOwner & "', '" &  TransType & "', '" & iPlayerClaimed & "', '" &  Cost & "')"
				objConn.Execute strSQL
					
		 		strSQL = "DELETE from tblWaivers where PID_Claimed =" & iPlayerClaimed & " and OwnerID = " & iOwner & " ;"
				objConn.Execute strSQL
					
			else
				'**********************************************************
				'UPDATE TO PLAYER TABLE for player being added.
				'**********************************************************
				strSQL = "update tblPlayers SET playerStatus = 'O', OwnerId = " & iOwner & ", pendingwaiver = 0, clearwaiverdate = null WHERE tblPlayers.PID = " & iPlayerClaimed & ";"
				objConn.Execute strSQL			
	
				'******************************************************************
				'UPDATE TO OWNERS TABLE.  Update other owners waiver priorities first 
				'then set the current owner's waiver priority to 12.
				'******************************************************************
				strSQL ="update tblowners SET waiverpriority = waiverpriority - 1 WHERE waiverpriority > " & iPriority & ";"
				objConn.Execute strSQL
			
				if iPlayerWaived = 0 then
					strSQL ="update tblowners SET waiverpriority = 12, ActivePlayerCnt = ActivePlayerCnt + 1 WHERE ownerid = " & iOwner & ";"
					objConn.Execute strSQL
				else
					strSQL ="update tblowners SET waiverpriority = 12 WHERE ownerid = " & iOwner & ";"
					objConn.Execute strSQL						
						
					'**********************************************************
					'Update Player Table for Player being Waived
					'**********************************************************
					strSQL = "update tblPlayers SET playerStatus = 'W', OwnerId = 0, pendingwaiver = 0, clearwaiverdate = date() + 1 WHERE tblPlayers.PID = " & iPlayerWaived & ";"
					objConn.Execute strSQL						
						
					'**********************************************************
					'Player Released TRANSACTION
					'**********************************************************
					TransType = "Released"
					Cost = 0.00
					strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_ 
					iOwner & "', '" &  TransType & "', '" & iPlayerWaived & "', '" &  Cost & "')"
					objConn.Execute strSQL						
			    end if				
		    
				'**********************************************************
				'Player Signed TRANSACTION
				'**********************************************************
				TransType = "Signed off waivers"
				Cost = 2.00
				strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_ 
				iOwner & "', '" &  TransType & "', '" &  iPlayerClaimed  & "', '" &  Cost & "')"
				objConn.Execute strSQL
				
		  		objRejectWaivers.Open "SELECT * FROM tblWaivers where PID_Claimed =" & iPlayerClaimed & " and OwnerID <> " & iOwner & " ;" , objConn
				TransType = "Waiver pickup rejected"
				Cost = 0.00			
					
				While Not objRejectWaivers.EOF
	
					iRejOwner = objRejectWaivers.Fields("OwnerID").Value
					iWaivedReject = objRejectWaivers.Fields("pid_waived").Value
					iClaimedReject = objRejectWaivers.Fields("pid_claimed").Value
					
					strSQL ="insert into tblTransactions(OwnerID,TransType,PID,TransCost) values ('" &_ 
					iRejOwner & "', '" &  TransType & "', '" &  iClaimedReject & "', '" &  Cost & "')"
					objConn.Execute strSQL
	
					'Update PendingWaiver flag
					strSQL = "update tblPlayers SET pendingwaiver = 0 WHERE tblPlayers.PID = " & iWaivedReject & ";"
					objConn.Execute strSQL
						
	                objRejectWaivers.MoveNext
				Wend

				objRejectWaivers.Close		

				'*************************************************************************
				'Delete all entries from tblWaivers table where player_id = Player Claimed
				'*************************************************************************
   		 		strSQL = "DELETE from tblWaivers where PID_Claimed = " & iPlayerClaimed & ";"
				objConn.Execute strSQL
				
				'*************************************************************************
				'Delete any additional rows from the tblWaivers table for the player that 
				'was just waived.  This is necessary if the owner had the same player on 
				'multiple waivers.
				'*************************************************************************
   		 		strSQL = "DELETE from tblWaivers where PID_Waived = " & iPlayerWaived & ";"
				objConn.Execute strSQL
			end if	    			
				
			'**********************************************************
			'Close the Query and Open it again to see if any rows remain
			'**********************************************************
			ObjRsTxns.Close		
			objRSTxns.Open		"SELECT * FROM qryUpdatewaiver ", objConn,3,3,1
    		w_action = objRSTxns.Recordcount
				
		wend
			
		ObjRsTxns.Close
			
		'********************************************************************
		'Make players Free whose clearwaiver Date is less then Date()
		'and Set Rental Players back Free
		'*********************************************************************			
		strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'F', OwnerId = 0, clearwaiverdate = null " & _
		         "WHERE clearwaiverdate < now() and playerStatus = 'W'"
		         
		objConn.Execute strSQL						
	   			
		strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'W', OwnerId = 0, rentalplayer = No, clearwaiverdate = date() + 1 " & _
		         "WHERE rentalplayer = Yes"
 
		objConn.Execute strSQL			
		
		'********************************************************************
		'Set the time for the next pendingwaiversall run.  If tomorrow is a game
		'day, then set the pendingwaivers date to run 6 hours before cutofftime.  
		'Note that that code subtracts 5 hours from the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a game day then set the nextrun date to be 1:00 PM EST.
		'*********************************************************************			
		objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", objConn,3,3,1

		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftime").Value - 5/24
		else
			dnextrun = date() + 1 + (13/24)
		end if

		strSQL = "update tbltimedEvents " & _
			     "SET lastrun = now(), nextrun = '"&dnextrun&"' " & _
		         "WHERE event = 'pendingwaiversall' "				
		         
		objConn.Execute strSQL					

	   	objNextRun.Close

	end if
	
	objRSwaivers.Close	

	'*************************************************
	'Run setwaivers event if it hasn't been run today.
	'*************************************************
	objRSwaivers.Open "SELECT * FROM tbltimedEvents " & _
	                  "where event = 'setwaiversall' and nextrun < now() ", objConn,3,3,1
	                  
	if  objRSwaivers.Recordcount > 0 then
		'Response.Write "Process setwaiverall logic .<br>"
		
		strSQL = "update tblPlayers " & _
		         "SET playerStatus = 'W', OwnerId = 0, RentalPlayer = 0, clearwaiverdate = date() + 1 " & _
		         "WHERE playerStatus = 'F' or RentalPlayer = 1 "
		objConn.Execute strSQL					


		'********************************************************************
		'Set the time for the next setwaiversall run.  If tomorrow is a game
		'day, then set the pendingwaivers date = to the cutofftime.  
		'Note that that code subtracts adds 1 hour to the time because the times
		'in the database are CST but the server is hosted on EST.  If tomorrow is
		'not a game day then set the nextrun date to be 7:00 PM EST.
		'*********************************************************************			
		objNextRun.Open	"SELECT * FROM qryGamedeadlines where gameday = (date() + 1) ", objConn,3,3,1

		if objNextRun.Recordcount > 0 then
			dnextrun = objNextRun.Fields("cutofftime").Value + 1/24
		else
			dnextrun = date() + 1 + (19/24)
		end if

		strSQL = "update tbltimedEvents " & _
			     "SET lastrun = now(), nextrun = '"&dnextrun&"' " & _
		         "WHERE event = 'setwaiversall' "				
		         
		objConn.Execute strSQL					

	   	objNextRun.Close

	end if
	objRSwaivers.Close	

	
	objRSgames.Open "qryGameDeadLines", objConn

%>
<html>
<head>
<title>home.asp</title>
<link rel="stylesheet" type="text/css" href="nav.css">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<base target="_self">
<style>
<!--
div {font-size: 13px; font-family: arial, helvetica, sans-serif; color: #000000;}	
.SLTables1 {font-size: 10px; font-family: verdana, arial, helvetica, sans-serif;}
.bg2 {background-color: #ededed;}
.bg1 {font-weight: bold; background-color: #b9b9b9;}
h1
	{margin-bottom:.0001pt;
	page-break-after:avoid;
	font-size:12.0pt;
	font-family:"Times New Roman";
	font-weight:normal; margin-left:0in; margin-right:0in; margin-top:0in}
-->
</style>
</head>
<body bgcolor="FFFFF7">

<div align="center">
    <center>
    <p></p>
    </center>
</div>
<div align="center">
    <center>
        <table border="1" cellpadding="2" cellspacing="0" width="699" height="2488">
          <tr>
            <th style="background-color: #FFFFFF" width="26" height="21">
            <font color="#800080" size="4"><b>RK</b></font></th>
            <th style="background-color: #FFFFFF" width="659" height="21" colspan="5">
            Team Analysis,<font color="#FFFF00"> </font><font color="#008000">Championship 
            Possibilities </font>and<font color="#FFFF00"> </font><font color="#FF0000">
            Needs</font><font color="#FFFF00"> </font><font color="#800080"><i>
            <font size="2">from Charles &amp; Rudi...</font></i></font></th>
          </tr>
          <tr>
            <td align="center" bgcolor="#008000" colspan="6" height="22" width="691">
            <font size="4" color="#FFFFFF">THE G.O.A.T. (18-8) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="476" rowspan="6">
            <font size="4"><b>1</b></font></td>
            <td width="127" height="20" align="center" bgcolor="#000099"><b>
			<font color="#FFFF00" size="2">STUD</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000099"><b>
			<font color="#FFFF00" size="2">STUD</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000099"><b>
			<font color="#FFFF00" size="2">STUD</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000099"><b>
			<font color="#FFFF00" size="2">OVERACHIEVER</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000099"><b>
			<font color="#FFFF00" size="2">UNDERACHIEVER</font></b></td>
          </tr>
          <tr>
            <td width="127" height="206">
			<p align="center">
			<img src="images/act_kobe_bryant.jpg" width="127" height="200"></td>
            <td width="127" height="206">
			<p align="center">
			<img src="images/act_chris_webber.jpg" width="127" height="200"></td>
            <td width="127" height="206">
			<p align="center">
			<img src="images/act_zach_randolph.jpg" width="127" height="200"></td>
            <td width="127" height="206">
			<p align="center">
			<img src="images/act_nazr_mohammed.jpg" width="127" height="200"></td>
            <td width="127" height="206">
			<p align="center">
			<img src="images/act_donyell_marshall.jpg" width="127" height="200"></td>
          </tr>
          <tr>
            <td width="127" height="20" align="center" bgcolor="#000000"><b>
			<font color="#FFFF00" size="2">43</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000000"><b>
			<font color="#FFFF00" size="2">37</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000000"><b>
			<font color="#FFFF00" size="2">35</font></b></td>
            <td width="127" height="20" align="center" bgcolor="#000000">
			<font size="2" color="#FFFF00"><b>25</b></font></td>
            <td width="127" height="20" align="center" bgcolor="#000000">
			<font size="2" color="#FFFF00"><b>21</b></font></td>
          </tr>
          <tr>
            <td width="659" height="134" colspan="5"><font size="2">
              <p>The G.O.A.T., what a name. First off despite a shaky schedule 
				in the month of December, GOAT has managed to hang on to first 
				place. We will see if this continues when Webber misses 7 of the 
				last 8 games to end the month. Very Strong trio of Kobe, Webber, 
				Zach. Bobby Simmons has been a pleasant surprise as has Nazi 
				Mohammad. Mohammad, a waiver pickup has been highly sought 
				after, but no reason to think he will be moved to due the spotty 
				schedule of Webber and Randolph. Crawford days as a big time 
				contributor are numbered due to the return of Alan Houston. Troy 
				Murphy has the yo-yo syndrome and can't be counted on, and the 
				rest of the bunch are decent but nothing to right home about.</p>
              </font></td>
          </tr>
          <tr>
            <td width="659" height="20" colspan="5"><font color="#008000">WILL 
			MAKE THE PLAYOFFS!</font></td>
          </tr>
          <tr>
            <td width="659" height="76" colspan="5"><font color="#FF0000">To 
			sure up a playoff team. With Kobe's playoff schedule, Webber's 
			health concerns, and Zach only available 3 times in the first round, 
			the time may be right to make a trade. Webber has proved he can post 
			great numbers despite lacking the ability to explode off his 
			repaired knee. His value is at its peak. Many teams could use center 
			help and may be willing to gamble on Webber. If Nazi is all that, 
			the time may be now to move Webber is now. </font></td>
          </tr>
          <tr bgcolor="#000099">
            <td height="28" colspan="6" align="center" valign="center" width="691" bgcolor="#000080">
            <font size="4" color="#FFFFFF">The Chill Factor (24-12) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="440" rowspan="3"><b><font size="4">
            2</font></b></td>
            <td bgcolor="#ffffff" width="659" height="289" colspan="5"><font size="2">
              <p>Duncan, Brad Miller, Peja, Kobe, and Francis, WOW! Last year everyone 
              would have just quit and gave this team the championship! Guess what, this 
              is a new season and questions surround all the aforementioned players but 
              Duncan (<i>second to KG only!</i>). Webber, will be returning this month and 
              he will have an impact on Peja and Brad as they try to work him in the 
              offense. Miller will take more of a hit than Peja I predict. Kobe has not 
              been Kobe which is to be expected with the legal troubles and additions of&nbsp; 
              Malone and Payton. Francis is struggling to play in Van Gundy's center 
              oriented offense causing his numbers to drop across the board. Rashard Lewis 
              is a valuable sixth man that will continue to get plenty of time as a 
              starter. Nothing else to talk about on the roster as this team is 6 deep! 
              Losing Harping to possible surgery may be a huge blow!</p>
              </font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="70" colspan="5"><font color="#008000">This team 
            has enough talent to win it all! Duncan, Peja, and Brad play 6xs in the first 
            round, Lewis 4, Kobe 3 and Francis 2. Others in the division are banking on 
            Webber slowing down Peja and Miller by then, leaving Duncan and as the anchor 
            to lead them. Kobe and Francis playing a combined 5 games may also cause this 
            team to falter. The second round Duncan only plays 2xs, but the others play 
            more and are capable of picking up the slack! </font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="70" colspan="5"><font color="#FF0000">Acquire 
            some guard depth for the first round and maybe even split the Sacramento 
            combo. Miller may need to stay because <b>Duncan</b> has a stretch where he 
            misses <b>9 of 12</b> games. We all know this owner has been reluctant to move 
            Peja all season. Don't bank on <b>Peja</b> going anywhere. The Chill Factor's 
            roster may be set for the run to the playoffs. The questions is will this be 
            enough for owner Cliff Fox who is salivating to win it all.</font></td>
          </tr>

          <tr>
            <td align="center" bgcolor="#000080" colspan="6" valign="center" height="22" width="691">
            <font size="4" color="#FFFFFF">Ballaticians 2.1.1.(25-11) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="244" rowspan="3"><font size="4"><b>3</b></font></td>
            <td width="659" height="144" colspan="5"><font size="2">
              <p>Season is now tied to the knees of Chris Webber as they traded a healthy 
              Shawn Marion (37 barps) for an oft-injured Star in Webber. Now Shaq comes up 
              lame. The injury does not seem to be serious, but this is Shaq! Ask Kobe 
              about Shaq and his willingness to play with pain. Team has been carried by 
              Baron Davis who has been fabulous this season. It's a good thing field goal 
              percentage is not factoring into our scoring system. When healthy, O'Neal, 
              Webber, Boozer, Richardson, and Davis take a back seat to no one, but 
              questions still remain. What will Webber be when he returns and how long 
              will it take for him to get in shape. Also, this is Webber so another injury 
              may occur. What will happen to Baron's numbers when Mashburn returns? Time 
              will tell as this team may fall from the top seed due to waiting in Webber.</p>
              </font></td>
          </tr>

          <tr>
            <td width="659" height="70" colspan="5"><font color="#008000">If Healthy will be a tough 
            first round draw! Shaq and Baron are only there 3xs in the first round, so 
            Webber, Kurt Thomas, Boozer, Richardson will have to be sharp and bring their 
            A games for success. Everyone has some weakness, and that is what this team is 
            banking on! They will have a cakewalk in the second round if this teams 
            advances to the conference finals. A serious contender to win their second 
            title if Webber steps up for once and leads instead of following! </font></td>
          </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#FF0000">A healthy <b>Webber</b>, and 
            a healthy <b>Webber</b>, and a healthy <b>Webber</b>!</font></td>
          </tr>

            <tr bgcolor="#000099">
                <td height="22" colspan="6" align="center" width="691" bgcolor="#008000">
            <font size="4" color="#FFFFFF"> GFOS (20-16)</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="284" rowspan="3"><font size="4"><b>4</b></font></td>
            <td width="659" height="128" colspan="5"><font size="2">
              <p>Guess whose back, Suga's back, Guess whose back, guess whose back, guess 
              whose back! Ray Allen has returned and has some friends with him. Rasheed 
              Wallace, Ray Allen, Antoine Walker, Jason Kidd, Andre Miller, Drew Gooden, 
              Rip Hamilton, Harry (K. Hinrich) Potter, Eddie Curry, and Chris Bosh make 
              this one of the deepest teams in the East. Everyone mentioned has been over 
              38 barps multiple times this season with the exception of Curry. With that 
              being said, what does that really mean. Nothing! Trading Pierce, Stoudemire 
              and Payton for Kidd, Wallace and Curry and the return of Ray Allen has this 
              team thinking finals. Who knows as this owner does not sit still and will 
              continue to make changes until the trading deadline.</p>
              </font></td>
          </tr>

          <tr>
            <td width="659" height="112" colspan="5"><font color="#008000">This team has positioned 
            itself to advance in the playoffs. Banking on Wallace though may not be a good 
            idea as he has a tendency to be very inconsistent. That being said, Wallace 
            can dominate anyone if his mind is right. Kidd, Allen, will have to be on 
            their games as they will be called upon to lead this team. I think their are 
            capable, but time will tell. Eddie Curry needs to get in shape, and there is 
            enough time for that to happen, Bosh needs to get his second wind also for 
            this team to be successful. When Gooden cracks Orlando's starting lineup (<b><i>it 
            will happen</i></b>) he should be lights out. Chances are very good this team 
            will advance to the conference finals. Beyond that is anyone's guess.</font></td>
          </tr>

          <tr>
            <td width="659" height="28" colspan="5"><font color="#FF0000">A patience owner as this 
            team seems to have decent depth at every position! Gooden inserted into the 
            starting lineup. Acquiring a better starting center would be nice also!</font></td>
          </tr>

          <tr bgcolor="#008000">
            <td align="center" colspan="6" height="22" width="691" bgcolor="#000080">
            <font size="4" color="#FFFFFF"> Nike Running Rebels (22-14) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="228" rowspan="3"><font size="4"><b>5</b></font></td>
            <td bgcolor="#ffffff" width="659" height="160" colspan="5"><font size="2">
              <p>Gasol, Ak-47, Abdur-Rahim, the newly acquired Cassell, Rose, Larry 
              Hughes, Ilgauskas, and Mashburn (<i><b>injured</b></i>) form a nice nucleus. 
              Lately this team has been torn on who to play because a lot of the players 
              are the same and none have stepped up to claim starting spots. Gasol should 
              be a nightly stud, but Hubie Brown want allow that to happen. Ilgauskas is 
              having his minutes jerked&nbsp; around by Silas, whose only concern is 
              getting Lebron ROY! Has Ak-47 run out of gas. Playing all out on a nightly 
              basis as well as giving away pounds (<i><b>starts at the 4</b></i>) may 
              start to take a toll on Kirilenko. Getting Cassell for Donyell Marshall was 
              a great move as Rose is on this roster as well. Should make the playoffs, 
              but needs Mashburn to come back hitting the ground running. Definitely has 
              talent but needs to set on a starting 5 rotation and live and die with it.</p>
              </font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="28" colspan="5"><font color="#008000">This team 
            has what others lack, playoff coverage due to its depth. Needs another top 
            notch player to be a serious contender. Also needs NBA coaches to let his 
            players play!</font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="24" colspan="5"><font color="#FF0000">Needs to 
            package talent and acquire a big name to combat the big guns in this division.</font></td>
          </tr>

          <tr bgcolor="#008000">
            <td height="22" colspan="6" align="center" width="691" bgcolor="#000080">
            <font size="4" color="#FFFFFF">King of the Hill (19-17) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="94" rowspan="3"><font size="4"><b>6</b></font></td>
            <td bgcolor="#ffffff" width="659" height="80" colspan="5"><font size="2">
              <p>Zach, Arenas, and Odom are the heart of this team. The newly acquired D. 
              Marshall will provide help but for how long is the question. He has a 
              tendency to fade down the stretch. Yao has been somewhat a disappointment 
              but still can provide huge games. Jason Terry is a good player, on a 
              horrible NBA team, who&nbsp; can dominate for games at a time. He also can 
              be just as horrible. Will Troy Murphy get into the starting lineup? Will 
              Jerome Williams remain in the starting lineup? Will Gilbert find out what is 
              eating him? Will Kwame Brown and C. Butler figure it out? Will Odom's 
              fragile ankles hold up? This team has a lot of questions and will struggle 
              if the above questions are not answered favorably. Zach Randolph will have 
              to remain solid, Arenas needs to get over his injuries and his self and 
              Marshall, Odom and Yao need to come up huge in the second half to keep this 
              team playoff aspirations alive.</p>
              </font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="7" colspan="5"><font color="#008000">Has enough 
            to win a round but not serious contenders for the title.</font></td>
          </tr>

          <tr>
            <td bgcolor="#ffffff" width="659" height="7" colspan="5"><font color="#FF0000">Needs 
            another big name&nbsp; to combat the duos of <b>Pierce/Brand, Shaq/Baron</b>, 
            and <b>Duncan/Peja</b>. Odom has been lights out, but we all know <b>Odom</b> 
            has never played a whole season. <b>Arenas</b> was that guy earlier in the 
            year but he has struggled since returning form his abdomen injury.</font></td>
          </tr>

          <tr bgcolor="#000099">
            <td align="center" colspan="6" height="22" width="691" bgcolor="#008000">
            <font size="4" color="#FFFFFF">Missing Sugar Ray (18-18)</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="92" rowspan="3"><font size="4"><b>7</b></font></td>
            <td width="659" height="64" colspan="5"><font size="2">
              <p>Inconsistency is the name of the game for the Defending Champions. 
              Centers are inconsistent! Forwards are inconsistent! Guards are 
              inconsistent! Camby can dominate a game with blocks and rebounds but he has 
              been banged up as of late. Howard has been a strange case to figure out 
              because he is 6-10 and can't get 5 rebounds on most nights. Antonio Davis 
              also seems lost! Carter, Richardson, Jones, Marbury, Houston have talent but 
              do not play well together on IGBL game nights. Carter and Marbury have to 
              step up and lead this team or they may be on the outside looking in. 
              Prediction: Vince will begin to become more assertive and have a big second 
              half! This will be needed or they may be headed to missing the playoffs 
              after winning it all last year. </p>
              </font>
            </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#008000">Has coverage but players 
            need to become more dependable as their inconsistencies make for to many 
            decisions on game nights. </font>
          </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#FF0000">Package 2 players for a <b>
            BIG TIME</b> performer! This want happen as this is the most <b>stubborn</b> 
            owner in the league! </font>
          </tr>

            <tr bgcolor="#000099">
                <td align="center" colspan="6" height="22" bgcolor="#008000" width="691">
            <font size="4" color="#FFFFFF">Playground Legends (17-19) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="108" rowspan="3"><font size="4"><b>8</b></font></td>
            <td width="659" height="80" colspan="5"><font size="2">
              <p>T-Mac and Radio are the pulse of this team. Richard Jefferson has been 
              playing well as of late which is needed for this team. Billups has been a 
              bust of late, but was playing well when he was first acquired for Ak-47 and 
              has the ability to be good again. The center situations is suspect, but so 
              is everyone else's in the east minus Dirk and Big Ben so that may not be a 
              problem. Chandler is said to be back in late January but who knows as back 
              injuries flare up from time to time and he was taking 8 pills daily when he 
              was playing. A healty return of Chandler is what this team needs. Glen Robinson has been hurt as has not looked good all year. 
              Overall a team that goes far as T-Mac and Radio will take them/ T-Mac has 
              had numerous injuries and it will be interesting to see what happens as the 
              Magic's Season goes further down the drain. Will they continue to play 
              T-Mac, or shut him down and look at their younger players?</p>
              </font></td>
          </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#008000">If the playoffs are made 
            they will be exited quickly as Radio and T-Mac both play 3 games each in the 
            first round and only one of those games is together. </font></td>
          </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#FF0000">Needs to acquire more talent 
            but these team's reputation is to Not Make any Trades out of fear of making 
            the wrong move. Since the <b>Ak-47</b> for <b>Billups</b> trade, no other 
            trade has been made and help is definitely needed. Will the lack of movement 
            allow the Runners the opportunity catch the Legends?</font></td>
          </tr>

          <tr bgcolor="#008000">
            <td align="center" colspan="6" height="22" bgcolor="#000080" width="691">
            <font size="4" color="#FFFFFF">Devastation Inc (15-21) </font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="96" rowspan="2"><font size="4"><b>9</b></font></td>
            <td width="659" height="48" colspan="5"><font size="2">
              <p>Brand, Pierce, Crawford and Artest are the cornerstones of this franchise.
                On paper, this team has the talent to make a playoff run and I predict 
              they will make the playoffs (<i><b>barring injury</b></i>). Stoudemire is 
              due back this month and when he gets back into the flow will help this team. 
              Bibby and Harrington are key reserves that have the capability to have big 
              games are key reserves. Payton may be huge if Malone and O'Neal miss a lot 
              of time due to injury. </p>
              </font></td>
          </tr>

          <tr>
            <td width="659" height="48" colspan="5"><font color="#FF0000">Trades have been made to set 
            this team in upward motion.&nbsp; Pierce needs to play well for a lot was 
            given to acquire his services.</font></td>
          </tr>

          <tr bgcolor="#008000">
            <td align="center" colspan="6" height="22" width="691">
            <font size="4" color="#FFFFFF">The Runners (12-24)</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" width="26" height="62" rowspan="2"><font size="4"><b>10</b></font></td>
            <td width="659" height="48" colspan="5"><font size="2">
              <p>Arguably the deepest team in the East: Dirk, Dampier, Marion, Malone, Van 
              Horn, James, Spreewell, and Mobley provide this team with enough talent to 
              make the playoffs. Some questioned the moves of trading Webber for Marion 
              and Kobe/Francis for Dirk/Mobley but it seemed this team new what it was 
              doing. Dirk's ankles are a question mark, but when healthy, he is a stud 
              evidenced by the two 48 barp performances put up last weekend. Set to make a 
              playoff run, the Runners need to put on the Nikes and start the chase of 
              teams that are above them. The Runners may benefit from the owners above 
              them, for they are not big on making trades.</p>
              </font></td>
          </tr>

          <tr>
            <td width="659" height="14" colspan="5"><font color="#FF0000">Be Patient and see if the 
            flurry of moves made last week pay off!</font></td>
          </tr>

          <tr bgcolor="#000099">
            <td align="center" colspan="6" height="22" bgcolor="#000080" width="691">
            <font size="4" color="#FFFFFF">Big Dogs (10-26)</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="96" rowspan="2">
            <font size="4"><b>11</b></font></td>
            <td width="659" height="64" colspan="5"><font size="2">Injuries, Injuries, Injuries has 
            cost this team. It is not lack of talent, if is just bad breaks. Injuries to 
            Iverson, Wade, Finley, Stackhouse, and Tim Thomas has just been to much for 
            this team to overcome. Playing in the West makes a playoff shot a lot more 
            difficult than if the Dogs were in the East or the Least as it has been 
            called. There have been collapses in previous years, and a collapse by others 
            and a hot streak by the Dogs is needed for them to get in the race. There is 
            36 games to go so anything is possible.</font></td>
          </tr>

          <tr>
            <td width="659" height="32" colspan="5"><font size="2" color="#FF0000">Trade for health as 
            JO and Iverson will both command a lot for the challenging for playoff spots!</font></td>
          </tr>

          <tr bgcolor="#000099">
            <td align="center" colspan="6" height="22" bgcolor="#008000" width="691">
            <font size="4" color="#FFFFFF">D-Town Bombers (9-27)</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#ce9c00" valign="center" width="26" height="147" rowspan="2">
            <font size="4"><b>12</b></font></td>
            <td width="659" height="98" colspan="5">What, When, How? What in the hell happened here! 
            When is Carmelo going to get traded? How in the hell do you allow yourself to 
            be in this position? Do you believe in beginners luck? After a successful 
            rookie season the Bombers have hit hard times. From a questionable auction, a 
            questionable draft and questionable trades, this team finds itself in Big 
            Trouble. I do admire the team for trying but they need to get on the stick 
            quick! There is some talent here but over estimation of its talent&nbsp; and 
            underestimations of other teams talent has been this team's biggest problem!
            <font size="2">There is 36 games to go so anything is possible.</font></td>
          </tr>

          <tr>
            <td width="659" height="41" colspan="5"><font color="#FF0000">Trade Redd or Carmelo and 
            see if 2 players can be acquired. They both have great playoff schedules early 
            and can be helpful to others who feel they are going to make the playoffs. 
            Holding them makes no sense if help can be acquired. </font></td>
          </tr>

      </table>
  </center>
</div>

</body>
</html>