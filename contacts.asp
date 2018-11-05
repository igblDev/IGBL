<!-- #include file="adovbs.inc" -->
<!--#include virtual="Common/IGBLStandard.inc"-->
<!DOCTYPE html>
<html>
<head>
<!--#include virtual="Common/pageHeader.inc"-->
<!-- Bootstrap core CSS -->
<!--#include virtual="Common/bootstrap.inc"-->
<!-- Application CSS -->
<link href="css/appblack.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<style>
th {
    vertical-align: middle;
}
td {
    vertical-align: middle;
}
black {
		color: black;
}
.panel-override {
  background-color:white;
	color:black;
}

</style>
</head>
<body>
<%
On Error Resume Next
Session("FP_OldCodePage") = Session.CodePage
Session("FP_OldLCID") = Session.LCID
Session.CodePage = 1252
Err.Clear

strErrorUrl = ""



   	Dim objConn,objRS
		Set objConn= Server.CreateObject("ADODB.Connection")
		Set objRS  = Server.CreateObject("ADODB.RecordSet")

		objConn.Open Application("lineupstest_ConnectionString")

    objConn.Open 	"Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=lineupstest.mdb;" & _
									"Persist Security Info=False"


	'objRS.Open "qryOwners", objConn
	objRS.Open "Select * from tblOwners where ownerID <> 99 order by TeamName", objConn
	
%>
<!--#include virtual="Common/session.inc"-->
<!--#include virtual="Common/headerMain.inc"-->
<div class="container">
	<div class="row">
		<div class="col-md-12 col-sm-12 col-xs-12">
			<div class="alert alert-banner">
				<strong><i class="fas fa-users"></i>&nbsp;Contact Owner(S)</strong>
			</div>
		</div>
	</div>
</div>

<div class="container">
  <div class="row">
    <div class="col-sm-12 col-md-12">
      <div class="panel panel-override">
        <div class="panel-body">
          <table class="table table-custom-black  	table-bordered table-striped table-condensed">
            <%
						While Not objRS.EOF
						%>
            <tr>
              <th colspan="2"><h4><%=objRS.Fields("TeamName").Value%></h4><img class="img-responsive" style="max-width:100%;height:auto;margin:0px auto;display:block;border: #111;border-width: thin;border-style: solid;" src="<%=objRS.Fields("teamlogo").Value%>"></th>
            </tr>
            <tr>
              <td><span style="color:blue;"><i class="fas fa-user"></i>&nbsp;<%=objRS.Fields("OwnerName").Value%>&nbsp;|&nbsp;<%=objRS.Fields("shortName").Value%></td>
            </tr>
						<tr>
              <td><span style="color:green;"><i class="fas fa-usd-circle"></i>&nbsp;Waiver Balance&nbsp;<%=FormatCurrency(objRS.Fields("waiverBal").Value)%></td>
            </tr>
						<tr>
              <td><span style="color:red;"><i class="fas fa-mobile-alt"></i></span>&nbsp;<%=objRS.Fields("CellPhone").Value%></td>
            </tr>
            <tr>
              <td><span style="color:purple;"><i class="fa fa-envelope" aria-hidden="true"></i></span>&nbsp;<%=objRS.Fields("homeEmail").Value%></td>
            </tr>
            <%
					 objRS.MoveNext
					 Wend
						%>
          </table>
        </div>
      </div>
    </div>
  </div>
</div>
<%
   objRS.Close
   objConn .Close
   Set objRS = Nothing
   Set objConn = Nothing
   Session.CodePage = Session("FP_OldCodePage")
   Session.LCID = Session("FP_OldLCID")
%>
</body>
</html>