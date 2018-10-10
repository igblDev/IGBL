						<%
						Function Forecast_Lineup (p_gameday, p_OwnerID, p_CenName, CEN_BARPS, p_For1Name, F1_BARPS, p_For2Name, F2_BARPS, p_Guard1Name, G1_BARPS, p_Guard2Name, G2_BARPS)
						
						p_Cen    = ""
						G1       = 0 
						G1_BARPS = 0
						G1_POS   = ""
						
						G2       = 0 
						G2_BARPS = 0
						G2_POS   = ""
						
						F1       = 0 
						F1_BARPS = 0
						F1_POS   = ""
						
						F2       = 0 
						F2_BARPS = 0
						F2_POS   = ""
						
						CEN      = 0  
						CEN_BARPS= 0
						CEN_POS  = ""
						 
						objRSPlayers.Open "SELECT * FROM qrymissinglineup " & _
															"WHERE gameday = #"&p_gameday&"# " & _
															"AND ownerID = "&p_OwnerID&" ", objConn,3,3,1
						 
						While Not objRSPlayers.EOF
								   
							if objRSPlayers.Fields("Pos").Value = "CEN" then
								if  CEN        = 0 then 
										CEN        = objRSPlayers.Fields("PID").Value
										CEN_BARPS  = objRSPlayers.Fields("barps").Value
										CEN_POS    = objRSPlayers.Fields("Pos").Value
										p_CenName  = objRSPlayers.Fields("lastName").Value
								elseif CEN_POS = "F-C" then
									if F1        = 0 then   'Move F-C currently assigned to Center to the open Forward 1.  Assign this player to Center
										F1         = CEN
										F1_BARPS   = CEN_BARPS
										F1_POS     = CEN_POS
										p_For1Name = p_CenName				 
										CEN        = objRSPlayers.Fields("PID").Value
										CEN_BARPS  = objRSPlayers.Fields("barps").Value
										CEN_POS    = objRSPlayers.Fields("Pos").Value
										p_CenName  = objRSPlayers.Fields("lastName").Value
										elseif F2  = 0 then  'Move F-C currently assigned to Center to the open Forward 2
										F2         = CEN
										F2_BARPS   = CEN_BARPS
										F2_POS     = CEN_POS 
										p_For2Name = p_CenName 
										CEN        = objRSPlayers.Fields("PID").Value
										CEN_BARPS  = objRSPlayers.Fields("barps").Value
										CEN_POS    = objRSPlayers.Fields("Pos").Value
										p_CenName  = objRSPlayers.Fields("lastName").Value       
									end if 
								end if								   
							elseif objRSPlayers.Fields("Pos").Value = "F-C" then
								if  CEN       = 0 then 
										CEN       = objRSPlayers.Fields("PID").Value
										CEN_BARPS = objRSPlayers.Fields("barps").Value
										CEN_POS   = objRSPlayers.Fields("Pos").Value
										p_CenName = left(objRSPlayers.Fields("lastName").Value,16)
								elseif F1     = 0 then   
										F1        = objRSPlayers.Fields("PID").Value
										F1_BARPS  = objRSPlayers.Fields("barps").Value
										F1_POS    = objRSPlayers.Fields("Pos").Value
										p_For1Name= left(objRSPlayers.Fields("lastName").Value,16)
								elseif F2     = 0 then
										F2        = objRSPlayers.Fields("PID").Value
										F2_BARPS  = objRSPlayers.Fields("barps").Value
										F2_POS    = objRSPlayers.Fields("Pos").Value
										p_For2Name= left(objRSPlayers.Fields("lastName").Value,16)
								end if				
							elseif objRSPlayers.Fields("Pos").Value = "FOR" then
								if  F1        = 0 then 
										F1        = objRSPlayers.Fields("PID").Value
										F1_BARPS  = objRSPlayers.Fields("barps").Value
										F1_POS    = objRSPlayers.Fields("Pos").Value
										p_For1Name= objRSPlayers.Fields("lastName").Value
								elseif F2     = 0 then
										F2        = objRSPlayers.Fields("PID").Value
										F2_BARPS  = objRSPlayers.Fields("barps").Value
										F2_POS    = objRSPlayers.Fields("Pos").Value 
										p_For2Name= left(objRSPlayers.Fields("lastName").Value,16)
								end if   								   
							elseif objRSPlayers.Fields("Pos").Value = "G-F" then
								if  G1          = 0 then
										G1          = objRSPlayers.Fields("PID").Value
										G1_BARPS    = objRSPlayers.Fields("barps").Value
										G1_POS      = objRSPlayers.Fields("Pos").Value 
										p_Guard1Name= left(objRSPlayers.Fields("lastName").Value,16)
								elseif G2       = 0 then
										G2          = objRSPlayers.Fields("PID").Value
										G2_BARPS    = objRSPlayers.Fields("barps").Value
										G2_POS      = objRSPlayers.Fields("Pos").Value 
										p_Guard2Name= left(objRSPlayers.Fields("lastName").Value,16)
								elseif F1       = 0 then   
										F1          = objRSPlayers.Fields("PID").Value
										F1_BARPS    = objRSPlayers.Fields("barps").Value
										F1_POS      = objRSPlayers.Fields("Pos").Value
										p_For1Name  = left(objRSPlayers.Fields("lastName").Value,16)
								elseif F2       = 0 then
										F2          = objRSPlayers.Fields("PID").Value
										F2_BARPS    = objRSPlayers.Fields("barps").Value
										F2_POS      = objRSPlayers.Fields("Pos").Value
										p_For2Name  = left(objRSPlayers.Fields("lastName").Value,16)
								end if												  
							else  'Guard Logic
								if  G1          = 0 then 
										G1          = objRSPlayers.Fields("PID").Value
										G1_BARPS    = objRSPlayers.Fields("barps").Value
										G1_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard1Name= left(objRSPlayers.Fields("lastName").Value,16)
								elseif G2       = 0 then
										G2          = objRSPlayers.Fields("PID").Value
										G2_BARPS    = objRSPlayers.Fields("barps").Value
										G2_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard2Name= left(objRSPlayers.Fields("lastName").Value,16)
								elseif G2_POS   = "G-F" then   
									if F1 = 0 then   'Move G-F currently assigned to Guard 2 to the open Forward 1.  Assign this player to Guard 2
										F1          = G2
										F1_BARPS    = G2_BARPS
										F1_POS      = G2_POS
										p_For1Name  = p_Guard2Name
										G2          = objRSPlayers.Fields("PID").Value
										G2_BARPS    = objRSPlayers.Fields("barps").Value
										G2_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard2Name= left(objRSPlayers.Fields("lastName").Value,16)
									elseif F2      = 0 then  'Move G-F currently assigned to Guard 2 to the open Forward 2.  Assign this player to Guard 2
										F2          = G2
										F2_BARPS    = G2_BARPS
										F2_POS      = G2_POS
										p_For2Name  = p_Guard2Name
										G2          = objRSPlayers.Fields("PID").Value
										G2_BARPS    = objRSPlayers.Fields("barps").Value
										G2_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard2Name= left(objRSPlayers.Fields("lastName").Value,16)
									end if										  
								elseif G1_POS = "G-F" then			
									if F1          = 0 then   'Move G-F currently assigned to Guard 1 to the open Forward 1.  Assign this player to Guard 1
										F1          = G1
										F1_BARPS    = G1_BARPS
										F1_POS      = G1_POS
										p_For1Name  = p_Guard1Name
										G1          = objRSPlayers.Fields("PID").Value
										G1_BARPS    = objRSPlayers.Fields("barps").Value
										G1_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard1Name= left(objRSPlayers.Fields("lastName").Value,16)
									elseif F2      = 0 then  'Move G-F currently assigned to Guard 1 to the open Forward 2.  Assign this player to Guard 1
										F2          = G1
										F2_BARPS    = G1_BARPS
										F2_POS      = G1_POS
										p_For2Name  = p_Guard1Name
										G1          = objRSPlayers.Fields("PID").Value
										G1_BARPS    = objRSPlayers.Fields("barps").Value
										G1_POS      = objRSPlayers.Fields("Pos").Value
										p_Guard1Name= left(objRSPlayers.Fields("lastName").Value,16)
									end if												  
								end if   	   
							end if			
							
						objRSPlayers.MoveNext
						Wend
						objRSPlayers.Close
							
							if F2_BARPS > F1_BARPS then
								w_Name      = p_For1Name
								p_For1Name  = p_For2Name
								p_For2Name  = w_Name
								
								w_Barps     = F1_BARPS
								F1_BARPS    = F2_BARPS
								F2_BARPS    = w_Barps
							end if
							
							if G2_BARPS > G1_BARPS then
								w_Name      = p_Guard1Name
								p_Guard1Name= p_Guard2Name
								p_Guard2Name= w_Name
								
								w_Barps     = G1_BARPS
								G1_BARPS    = G2_BARPS
								G2_BARPS    = w_Barps
							end if
						
					  End Function	
						%>
						