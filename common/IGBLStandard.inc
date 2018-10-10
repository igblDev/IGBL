<%
'*******************************************************************************
'* Script Name: IGBLStandard.inc
'* Description: Standard include file for IGBL applications.
'*******************************************************************************
'* Usage Key
'*   - Private (should only be used within Standard.inc)
'*   + Restricted (should only be used by other global include files)
'*   * Public (may be used universally)
'*
'* Errors
'*   - EnvironmentVariables : string
'*   - LogErrorByMail : boolean
'*   - LogErrorByEvent
'*   * RaiseGeneralError
'*   * RaiseSystemError
'*
'* Strings
'*   * IsBlank : string
'*
'* Registry Access
'*   - RegistryValue : string
'*   - GetRegistryValues
'*
'* Security
'*
'* Database Access
'*   + dbConnectWithRegistry
'*   + dbConnect
'*   + dbGetUserData
'*   + dbQuery : recordset object
'*   + dbCommand
'*   + dbBeginTransaction
'*   + dbCommitTransaction
'*   + dbRollbackTransaction
'*   + dbDisconnect
'*   * dbQuoted : string
'*   * dbDate : date
'*   * dbDefaultedDate : date
'*   * dbInteger : integer
'*   * dbDefaultedInteger : integer
'*   * dbSingle : single
'*   * dbDefaultedSingle : single
'*   * dbString : string
'*   * dbDefaultedString : string
'*
'* E-mail
'*   * gciEmailBodyWidth
'*   * IsValidEmailAddress : boolean
'*   * RegisteredEmailAddresses : string
'*   * SendEMail
'*
'* Parameters
'*   * GetLinkParameter
'*   * GetFormParameter
'*   * GetAnyParameter
'*   * AddLinkParameter
'*
'* Dates
'*   * gdtNow
'*   * gdToday
'*   * gsDBToday
'*   * DateToDBDate
'*   * DBDateToDate
'*   * StandardFormattedDate
'*   * StandardFormattedDateTime
'*
'* Formatting
'*   - RomanNumeralSubstring : string
'*   * FormattedRomanNumeral : string
'*   * WrappedText : string
'*
'* User Interface
'*******************************************************************************

	const gcbDebug = True
	const gcbProd  = False

'*******************************************************************************
'*******************************************************************************
'* Errors
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'* EnvironmentVariables
'*   Returns a text string containing the environment variables.
'*******************************************************************************
function EnvironmentVariables
	dim   lsEnvironment, lsItem

	'Retrieve standard environment variables.
	lsEnvironment = _
		"Standard" _
& vbNewLine &	"====================" _
& vbNewLine &	"AUTH_TYPE            " & Request.ServerVariables("AUTH_TYPE") _
& vbNewLine &	"AUTH_REALM           " & Request.ServerVariables("AUTH_REALM") _
& vbNewLine &	"CONTENT_FILE         " & Request.ServerVariables("CONTENT_FILE") _
& vbNewLine &	"CONTENT_LENGTH       " & Request.ServerVariables("CONTENT_LENGTH") _
& vbNewLine &	"CONTENT_TYPE         " & Request.ServerVariables("CONTENT_TYPE") _
& vbNewLine &	"DOCUMENT_ROOT        " & Request.ServerVariables("DOCUMENT_ROOT") _
& vbNewLine &	"GATEWAY_INTERFACE    " & Request.ServerVariables("GATEWAY_INTERFACE") _
& vbNewLine &	"HTTP_CONNECTION      " & Request.ServerVariables("HTTP_CONNECTION") _
& vbNewLine &	"HTTP_REFERER         " & Request.ServerVariables("HTTP_REFERER") _
& vbNewLine &	"HTTP_USER_AGENT      " & Request.ServerVariables("HTTP_USER_AGENT") _
& vbNewLine &	"LOGON_USER           " & Request.ServerVariables("REMOTE_USER") _
& vbNewLine &	"LOGON_PWD            " & Request.ServerVariables("REMOTE_PWD") _
& vbNewLine &	"PATH_INFO            " & Request.ServerVariables("PATH_INFO") _
& vbNewLine &	"PATH_TRANSLATED      " & Request.ServerVariables("PATH_TRANSLATED") _
& vbNewLine &	"QUERY_STRING         " & Request.ServerVariables("QUERY_STRING") _
& vbNewLine &	"REMOTE_ADDR          " & Request.ServerVariables("REMOTE_ADDR") _
& vbNewLine &	"REMOTE_HOST          " & Request.ServerVariables("REMOTE_HOST") _
& vbNewLine &	"REQUEST_METHOD       " & Request.ServerVariables("REQUEST_METHOD") _
& vbNewLine &	"SCRIPT_MAP           " & Request.ServerVariables("SCRIPT_MAP") _
& vbNewLine &	"SCRIPT_NAME          " & Request.ServerVariables("SCRIPT_NAME") _
& vbNewLine &	"SERVER_ADMIN         " & Request.ServerVariables("SERVER_ADMIN") _
& vbNewLine &	"SERVER_NAME          " & Request.ServerVariables("SERVER_NAME") _
& vbNewLine &	"SERVER_PORT          " & Request.ServerVariables("SERVER_PORT") _
& vbNewLine &	"SERVER_PORT_SECURE   " & Request.ServerVariables("SERVER_PORT_SECURE") _
& vbNewLine &	"SERVER_PROTOCOL      " & Request.ServerVariables("SERVER_PROTOCOL") _
& vbNewLine &	"SERVER_SOFTWARE      " & Request.ServerVariables("SERVER_SOFTWARE") _
& vbNewLine &	"URL                  " & Request.ServerVariables("URL")

	'Retrieve GET variables.
	lsEnvironment = lsEnvironment _
& vbNewLine _
& vbNewLine &	"GET" _
& vbNewLine &	"===================="
	for each lsItem in Request.QueryString
	  if InStr(lsItem, "Password") = 0 then
	    if Len(lsItem) <= 20 then
	      lsEnvironment = lsEnvironment _
& vbNewLine &	Left(lsItem & "                     ", 21) & Request.QueryString(lsItem)
	    else
	      lsEnvironment = lsEnvironment _
& vbNewLine &	lsItem & " " & Request.QueryString(lsItem)
	    end if
	  end if
	next

	'Retrieve POST variables.
	lsEnvironment = lsEnvironment _
& vbNewLine _
& vbNewLine &	"POST" _
& vbNewLine &	"===================="
	for each lsItem in Request.Form
	  if InStr(lsItem, "Password") = 0 then
	    if Len(lsItem) <= 20 then
	      lsEnvironment = lsEnvironment _
& vbNewLine &	Left(lsItem & "                     ", 21) & Request.Form(lsItem)
	    else
	      lsEnvironment = lsEnvironment _
& vbNewLine &	lsItem & " " & Request.Form(lsItem)
	    end if
	  end if
	next

	EnvironmentVariables = lsEnvironment
end function

'*******************************************************************************
'* RaiseGeneralError
'*   Raises a general error.  Does not return.
'*   Used for non-system errors (do not require administrator notification).
'*******************************************************************************
sub RaiseGeneralError _
	(byVal vsTitle, byVal vsMessage)
	'vsTitle             error title
	'vsMessage           error message

	'Display an error message and stop.
        Response.Clear
%>
<html>
<head><title><% = vsTitle %></title></head>
<body bgcolor="#FFFFFF" link="#003399" vlink="#003300" alink="#6699FF">
<div align="center"><center>
<table width="450" border="0" cellpadding="2" cellspacing="0">
 <tr <% = gsTTBG(0) %>>
  <td align="center"><% = gsTTF0(0) & vsTitle & gsTTF1(0) %></td>
 </tr><tr <% = gsTDBG(0,0) %>>
  <td><% = gsTDF0(0,0) & vsMessage & gsTDF1(0,0) %></td>
 </tr>
</table>
</center></div>
</body>
</html>
<%
        Response.End
end sub


'*******************************************************************************
'* RaiseSystemError
'*   Raises a system error.  Does not return.
'*   Used for system errors (require administrator notification).
'*******************************************************************************
sub RaiseSystemError _
	(byVal vsError)
	'vsError             error message

	dim   lsNumber, lsSource, lsDescription, lsEnvironment
	dim   loEventLog

	'Retrieve the error information to prevent overwrite.
	lsNumber = Err.Number
	lsSource = Err.Source
	lsDescription = Err.Description
	lsEnvironment = EnvironmentVariables

	'If debugging is in progress...
	if gcbDebug then
	  'Display an error message and stop.
%>
<hr>
<h1>Error Detected</h1>
<table border="0" cellpadding="0" cellspacing="4">
 <tr>
  <th align="left" valign="top">Symptom:</th>
  <td align="left" valign="top"><pre><% = vsError %></pre></td>
 </tr><tr>
  <th align="left" valign="top">Number:</th>
  <td align="left" valign="top"><% = lsNumber %></td>
 </tr><tr>
  <th align="left" valign="top">Source:</th>
  <td align="left" valign="top"><% = lsSource %></td>
 </tr><tr>
  <th align="left" valign="top">Description:</th>
  <td align="left" valign="top"><% = lsDescription %></td>
 </tr><tr>
  <th align="left" valign="top">Environment:</th>
  <td align="left" valign="top"><pre><% = lsEnvironment %></pre></td>
 </tr>
</table>
<hr>
<%
	  Response.End
	else '(debugging is not in progress)
	  'If an e-mail cannot be sent...
          if not LogErrorByMail(lsSource, lsNumber, vsError & vbNewLine & vbNewLine & lsDescription, lsEnvironment) then
            'Log the error in the event log.
            LogErrorByEvent lsSource, lsNumber, vsError & vbNewLine & vbNewLine & lsDescription, lsEnvironment
          end if

          'Display an error message and stop.
          Response.Clear
%>
<html>
<head><title>Unable to Service Request</title></head>
<body bgcolor="#FFFFFF" link="#003399" vlink="#003300" alink="#6699FF">
<div align="center"><center>
<table width="450" border="0" cellpadding="2" cellspacing="0">
 <tr <% = gsTTBG(0) %>>
  <td align="center"><% = gsTTF0(0) & "unable to service request" & gsTTF1(0) %></td>
 </tr><tr <% = gsTDBG(0,0) %>>
  <td><% = gsTDF0(0,0) & "An internal problem has occurred, been logged and will be forwarded to a TRAIN administrator.<br><br>Please try again later." & gsTDF1(0,0) %></td>
 </tr>
</table>
</center></div>
</body>
</html>
<%
          Response.End
	end if
end sub



'*******************************************************************************
'*******************************************************************************
'* Strings
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'* IsBlank
'*   Returns True if the input string is null or all whitespace, False otherwise.
'*******************************************************************************
function IsBlank _
	(byVal vsInput)

	if IsNull(vsInput) then
	  IsBlank = True
	elseif Trim(vsInput) = "" then
	  IsBlank = True
	else
	  IsBlank = False
	end if
end function



'*******************************************************************************
'*******************************************************************************
'* Security
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'*******************************************************************************
'* Registry Access
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'* RegistryValue
'*   Returns a value from the registry.
'*******************************************************************************
function RegistryValue _
	(byVal vsKey, byVal vsName)
	'vsKey               key of value to retrieve
	'vsName              name of value to retrieve

	dim   loRegistry, lsDomain, lsValue

	on error resume next

	'Retrieve the value.
	'set loRegistry = Server.CreateObject("IISSample.RegistryAccess")
	set loRegistry = g_oUtils
	'if Err.Number <> 0 then RaiseSystemError "Standard.inc / RegistryValue: Unable to create Registry object (" & vsKey & "\" & vsName & ")."
	'lsValue = loRegistry.Get("HKLM\" & vsKey & "\" & vsName)
	lsValue = loRegistry.regValueStringByName(vsKey, vsName)
	'if Err.Number <> 0 then RaiseSystemError "Standard.inc / RegistryValue: Unable to retrieve value (" & vsKey & "\" & vsName & ")."
	set loRegistry = nothing
	RegistryValue = lsValue
end function

'*******************************************************************************
'* GetRegistryValues
'*   Retrieves a registry value and parses it into its subcomponents.
'*   The retrieved components are placed in a dictionary which must be disposed
'*     of by the caller.
'*******************************************************************************
sub GetRegistryValues _
	(byVal vsKey, byVal vsValue, byRef rodComponents)

	dim   lasComponents, liComponent, lsComponent, liSeparator

	on error resume next

	'Create the dictionary object.
	set rodComponents = Server.CreateObject("Scripting.Dictionary")
	if Err.Number <> 0 then ErrorRaise "Standard.inc / GetRegistryValues: Unable to create Dictionary object."

	'Retrieve the value and split it into components.
	lasComponents = Split(RegistryValue(vsKey, vsValue), ";")
	for liComponent = 0 to UBound(lasComponents)
	  'Split each component into a key/value pair.
	  lsComponent = lasComponents(liComponent)
	  liSeparator = InStr(lsComponent, "=")
	  if liSeparator > 0 then
	    rodComponents.Add Left(lsComponent, liSeparator - 1), Right(lsComponent, Len(lsComponent) - liSeparator)
	  else
	    rodComponents.Add lasComponents(liComponent), ""
	  end if
	next
	set lasComponents = nothing
end sub

'*******************************************************************************
'*******************************************************************************
'* Database Access
'*******************************************************************************
'*******************************************************************************
	dim   gsUserID, gsFullName

'*******************************************************************************
'* dbQuoted
'*   Prepares a string value for use in a database query.
'*******************************************************************************
function dbQuoted _
	(byVal vsInput)
	'vsInput             string to prepare

	dbQuoted = "'" & Replace(vsInput, "'", "''") & "'"
end function


'*******************************************************************************
'* dbDate
'*   Returns a date value from a field.
'*******************************************************************************
function dbDate _
	(byVal vcInput, byVal vsFieldName)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve

	dim   ldValue

	on error resume next

	'Retrieve the value of the field.
	ldValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbDate: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(ldValue) then
	  'Return the null value.
	  dbDate = ldValue
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsDate(ldValue) then RaiseSystemError "Standard.inc / dbDate: Unable to convert value in " & vsFieldName & " to date format."

	  'Return the value as a date.
	  dbDate = CDate(ldValue)
	end if
end function


'*******************************************************************************
'* dbDefaultedDate
'*   Returns a date value from a field, with null values taking a specified
'*     default value.
'*******************************************************************************
function dbDefaultedDate _
	(byVal vcInput, byVal vsFieldName, byVal vdDefault)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve
	'vdDefault           default date value

	dim   ldValue

	on error resume next

	'Retrieve the value of the field.
	ldValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbDefaultedDate: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(ldValue) then
	  'Return the default value.
	  dbDefaultedDate = vdDefault
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsDate(ldValue) then RaiseSystemError "Standard.inc / dbDefaultedDate: Unable to convert value in " & vsFieldName & " to date format."

	  'Return the value as a date.
	  dbDefaultedDate = CDate(ldValue)
	end if
end function


'*******************************************************************************
'* dbInteger
'*   Returns a integer value from a field.
'*******************************************************************************
function dbInteger _
	(byVal vcInput, byVal vsFieldName)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve

	dim   liValue

	'Retrieve the value of the field.
	liValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbInteger: Unable to retrieve field " & vsFieldName & " from cursor."

	on error resume next

	'If the value is null...
	if IsNull(liValue) then
	  'Return the null value.
	  dbInteger = liValue
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsNumeric(CStr(liValue)) then RaiseSystemError "Standard.inc / dbInteger: Unable to convert value in " & vsFieldName & " to integer format."

	  'Return the value as a integer.
	  dbInteger = CLng(liValue)
	end if
end function


'*******************************************************************************
'* dbDefaultedInteger
'*   Returns a integer value from a field, with null values taking a specified
'*     default value.
'*******************************************************************************
function dbDefaultedInteger _
	(byVal vcInput, byVal vsFieldName, byVal viDefault)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve
	'viDefault           default integer value

	dim   liValue

	'Retrieve the value of the field.
	liValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbDefaultedInteger: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(liValue) then
	  'Return the default value.
	  dbDefaultedInteger = viDefault
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsNumeric(CStr(liValue)) then RaiseSystemError "Standard.inc / dbDefaultedInteger: Unable to convert value in " & vsFieldName & " to integer format."

	  'Return the value as a integer.
	  dbDefaultedInteger = CLng(liValue)
	end if
end function


'*******************************************************************************
'* dbSingle
'*   Returns a single-precision floating point value from a field.
'*******************************************************************************
function dbSingle _
	(byVal vcInput, byVal vsFieldName)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve

	dim   lsfValue

	'Retrieve the value of the field.
	lsfValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbSingle: Unable to retrieve field " & vsFieldName & " from cursor."

	on error resume next

	'If the value is null...
	if IsNull(lsfValue) then
	  'Return the null value.
	  dbSingle = lsfValue
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsNumeric(CStr(lsfValue)) then RaiseSystemError "Standard.inc / dbSingle: Unable to convert value in " & vsFieldName & " to single format."

	  'Return the value as a single.
	  dbSingle = CSng(lsfValue)
	end if
end function


'*******************************************************************************
'* dbDefaultedSingle
'*   Returns a single-precision floating point value from a field, with null
'*   values taking a specified default value.
'*******************************************************************************
function dbDefaultedSingle _
	(byVal vcInput, byVal vsFieldName, byVal vsfDefault)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve
	'vsfDefault          default single value

	dim   lsfValue

	'Retrieve the value of the field.
	lsfValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbDefaultedSingle: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(lsfValue) then
	  'Return the default value.
	  dbDefaultedSingle = vsfDefault
	else '(the value is not null)
	  'Verify that the value can be converted.
	  if not IsNumeric(CStr(lsfValue)) then RaiseSystemError "Standard.inc / dbDefaultedSingle: Unable to convert value in " & vsFieldName & " to single format."

	  'Return the value as a single.
	  dbDefaultedSingle = CSng(lsfValue)
	end if
end function


'*******************************************************************************
'* dbString
'*   Returns a string value from a field.
'*******************************************************************************
function dbString _
	(byVal vcInput, byVal vsFieldName)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve

	dim   lsValue

	on error resume next

	'Retrieve the value of the field.
	lsValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbString: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(lsValue) then
	  'Return the null value.
	  dbString = lsValue
	else '(the value is not null)
	  'Return the value as a string.
	  dbString = CStr(lsValue)
	end if
end function


'*******************************************************************************
'* dbDefaultedString
'*   Returns a string value from a field, with null values taking a specified
'*     default value.
'*******************************************************************************
function dbDefaultedString _
	(byVal vcInput, byVal vsFieldName, byVal vsDefault)
	'vcInput             cursor to retrieve value from
	'vsFieldName         name of field to retrieve
	'vsDefault           default string value

	dim   lsValue

	on error resume next

	'Retrieve the value of the field.
	lsValue = vcInput.Fields(vsFieldName)
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / dbDefaultedString: Unable to retrieve field " & vsFieldName & " from cursor."

	'If the value is null...
	if IsNull(lsValue) then
	  'Return the default value.
	  dbDefaultedString = vsDefault
	else '(the value is not null)
	  'Return the value as a string.
	  dbDefaultedString = CStr(lsValue)
	end if
end function

'*******************************************************************************
'*******************************************************************************
'* E-mail
'*******************************************************************************
'*******************************************************************************
	'Maximum width of the body of an E-mail message
	const gciEmailBodyWidth = 70

'*******************************************************************************
'* IsValidEmailAddress
'*   Returns True if the input appears to be a valid E-mail address.
'*   The format is: c{c}@c{c}.{{c}.}c{c}, where c represents any character
'*     except '.' and '@', and items in brackets may be repeated zero or more
'*     times.
'*******************************************************************************
function IsValidEmailAddress _
	(byVal vsAddress)
	'vsAddress           'address to validate

	dim   lasSplitAt, lasSplitDot

	lasSplitAt = Split(vsAddress, "@")
	if UBound(lasSplitAt) <> 1 then
	  IsValidEmailAddress = False
	elseif Len(lasSplitAt(0)) = 0 then
	  IsValidEmailAddress = False
	else
	  lasSplitDot = Split(lasSplitAt(1), ".")
	  if UBound(lasSplitDot) < 1 then
	    IsValidEmailAddress = False
	  elseif Len(lasSplitDot(0)) = 0 then
	    IsValidEmailAddress = False
	  elseif Len(lasSplitDot(UBound(lasSplitDot))) = 0 then
	    IsValidEmailAddress = False
	  else
	    IsValidEmailAddress = True
	  end if
	  set lasSplitDot = nothing
	end if
	set lasSplitAt = nothing
end function


'*******************************************************************************
'* SendEmail
'*   Returns a list of E-mail addresses from the registry.
'*******************************************************************************
sub SendEMail _
	(byVal vsFrom, byVal vsTo, byVal vsCC, byVal vsBCC, byVal vsSubject, byVal vsBody, byVal viFiles, byVal vasFileName)
	'vsFrom              from address
	'vsTo                to address list
	'vsCC                cc address list
	'vsBCC               bcc address list
	'vsSubject           subject of message
	'vsBody              body of message
	'viFiles             number of attachments
	'vasFileName         list of attachments

	dim   loNewMail, liFile

	on error resume next

	'Create the message.
	set loNewMail = Server.CreateObject("CDONTS.NewMail")
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / SendEMail: Unable to create NewMail object."
	loNewMail.From = vsFrom
	loNewMail.To = vsTo
	loNewMail.Cc = vsCC
	loNewMail.Bcc = vsBCC
	loNewMail.Subject = vsSubject
	loNewMail.Body = vsBody

	'Insert attachments.
	for liFile = 1 to viFiles
	  loNewMail.MailFormat = 0
	  loNewMail.AttachFile vasFileName(liFile - 1), , 1
	next

	'Send the message.
	loNewMail.Send
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / SendEMail: Unable to send message from " & vsFrom & " to " & vsTo & " regarding " & vsSubject & "."
	set loNewMail = nothing
end sub

'*******************************************************************************
'*******************************************************************************
'* Parameters
'*******************************************************************************
'*******************************************************************************

'*******************************************************************************
'* GetLinkParameter
'*   Retrieves a link parameter, setting nulls to "".
'*******************************************************************************
sub GetLinkParameter _
	(byVal vsName, byRef rsValue)
	'vsName              parameter name
	'rsValue             retrieved non-null value

	rsValue = Request.QueryString(vsName)
	if IsNull(rsValue) then rsValue = ""
end sub

'*******************************************************************************
'* GetFormParameter
'*   Retrieves a form parameter, setting nulls to "".
'*******************************************************************************
sub GetFormParameter _
	(byVal vsName, byRef rsValue)
	'vsName              parameter name
	'rsValue             retrieved non-null value

	rsValue = Request.Form(vsName)
	if IsNull(rsValue) then rsValue = ""
end sub

'*******************************************************************************
'* GetAnyParameter
'*   Retrieves a link/form/other parameter, setting nulls to "".
'*******************************************************************************
sub GetAnyParameter _
	(byVal vsName, byRef rsValue)
	'vsName              parameter name
	'rsValue             retrieved non-null value

	'If a link parameter exists with the specified name...
	if Request.QueryString(vsName).Count > 0 then
	  'Set the value to the link parameter.
	  rsValue = Request.QueryString(vsName)
	'If a form parameter exists with the specified name...
	elseif Request.Form(vsName).Count > 0 then
	  'Set the value to the form parameter.
	  rsValue = Request.Form(vsName)
	else '(no link or form parameters exist with the specified name)
	  'Set the value to the empty string.
	  rsValue = ""
	end if
end sub


'*******************************************************************************
'* AddLinkParameter
'*   Appends a link parameter to a URL.
'*******************************************************************************
sub AddLinkParameter _
	(byVal vsName, byVal vsValue, byRef rsURL)
	'vsName              parameter name
	'vsValue             value to add
	'rsURL               URL to add parameter to

	if InStr(rsURL, "?") then
	  rsURL = rsURL & "&" & vsName & "=" & Server.URLEncode(vsValue)
	else
	  rsURL = rsURL & "?" & vsName & "=" & Server.URLEncode(vsValue)
	end if
end sub


'*******************************************************************************
'*******************************************************************************
'* Dates
'*******************************************************************************
'*******************************************************************************
	dim   gdtNow         'current date/time in internal format
	gdtNow = Now

	dim   gdToday        'today's date in internal format
	gdToday = DateValue(CStr(gdtNow))

	dim   gsDBToday      'today's date in database format
	gsDBToday = DateToDBDate(gdToday)


'*******************************************************************************
'* DateToDBDate
'*   Returns a date converted from internal format to database format.
'*******************************************************************************
function DateToDBDate _
	(byVal vdInput)
	'vdInput             date to convert

	DateToDBDate = CStr(DatePart("yyyy", vdInput)) & "-" & Right("0" & CStr(DatePart("m", vdInput)), 2) & "-" & Right("0" & CStr(DatePart("d", vdInput)), 2)
end function


'*******************************************************************************
'* DBDateToDate
'*   Returns a date converted from database format to internal format.
'*******************************************************************************
function DBDateToDate _
	(byVal vsInput)
	'vsInput             date string to convert

	dim   lasDateComponents, ldResult

	lasDateComponents = Split(vsInput, "-")
	if UBound(lasDateComponents) <> 2 then RaiseSystemError "Standard.inc / DBDateToDate: Input date (" & vsInput & ") is not in proper format."
	ldResult = DateSerial(CInt(lasDateComponents(0)), CInt(lasDateComponents(1)), CInt(lasDateComponents(2)))
	if Err.Number <> 0 then RaiseSystemError "Standard.inc / DBDateToDate: Unable to convert date (" & vsInput & ")."
	set lasDateComponents = nothing
	DBDateToDate = ldResult
end function


'******************************************************************************
'* StandardFormattedDate
'*  Returns a date in the format ddmonyyyy.
'******************************************************************************
function StandardFormattedDate _
	(byVal vdInput)
	'vdInput             date to format

	StandardFormattedDate = Right("00" & CStr(Day(vdInput)), 2) & MonthName(Month(vdInput), True) & CStr(Year(vdInput))
end function


'******************************************************************************
'* StandardFormattedDateTime
'*  Returns a date in the format ddmonyyyy hh:mm.
'******************************************************************************
function StandardFormattedDateTime _
	(byVal vdtInput)
	'vdtInput            date/time to format

	StandardFormattedDateTime = _
	  Right("00" & CStr(Day(vdtInput)), 2) & MonthName(Month(vdtInput), True) & CStr(Year(vdtInput)) & "&nbsp;" _
	& Right("00" & CStr(Hour(vdtInput)), 2) & ":" & Right("00" & CStr(Minute(vdtInput)), 2)
end function

'*******************************************************************************
'*******************************************************************************
'* Formatting
'*******************************************************************************
'*******************************************************************************


'******************************************************************************
'* RomanNumeralSubstring
'*   Returns a roman numeral substring for the specified digit.
'******************************************************************************
function RomanNumeralPlace _
	(byVal viDigit, byVal vs10, byVal vs5, byVal vs1)
	'viDigit             digit to romanize (0..9)
	'vs10                numeral representing 10 at digit's original scale
	'vs5                 numeral representing 5 at digit's original scale
	'vs1                 numeral representing 1 at digit's original scale

	select case viDigit
	  case 0: RomanNumeralPlace = ""
	  case 1: RomanNumeralPlace = vs1
	  case 2: RomanNumeralPlace = vs1 & vs1
	  case 3: RomanNumeralPlace = vs1 & vs1 & vs1
	  case 4: RomanNumeralPlace = vs1 & vs5
	  case 5: RomanNumeralPlace = vs5
	  case 6: RomanNumeralPlace = vs5 & vs1
	  case 7: RomanNumeralPlace = vs5 & vs1 & vs1
	  case 8: RomanNumeralPlace = vs5 & vs1 & vs1 & vs1
	  case 9: RomanNumeralPlace = vs1 & vs10
	end select
end function


'******************************************************************************
'* FormattedRomanNumeral
'*  Returns the roman numeral equivalent of the specified number.
'******************************************************************************
function FormattedRomanNumeral _
	(byVal viInput)
	'viInput             number to format

	dim   lbIsNegative, lsOutput
	dim   liDigit

	'Ensure that the input is integral.
	viInput = Round(viInput, 0)

	'If the input is negative, record that fact and negate the input.
	lbIsNegative = (viInput < 0)
	if lbIsNegative then viInput = 0 - viInput

	'Initialize the output.
	lsOutput = ""

	'Append the thousands place.
	liDigit = viInput \ 1000
	lsOutput = lsOutput & String(liDigit, "M")
	viInput = viInput - (liDigit * 1000)

	'Append the hundreds place.
	liDigit = viInput \ 100
	lsOutput = lsOutput & RomanNumeralPlace(liDigit, "M", "D", "C")
	viInput = viInput - (liDigit * 100)

	'Append the tens place.
	liDigit = viInput \ 10
	lsOutput = lsOutput & RomanNumeralPlace(liDigit, "C", "L", "X")
	viInput = viInput - (liDigit * 10)

	'Append the ones place.
	liDigit = viInput \ 10
	lsOutput = lsOutput & RomanNumeralPlace(viInput, "X", "V", "I")

	'If the original input was negative, parenthesize the output.
	if lbIsNegative then lsOutput = "(" & lsOutput & ")"

	FormattedRomanNumeral = lsOutput
end function

'*******************************************************************************
'* WrappedText
'*   Returns a text string wrapped to the specified width.
'*******************************************************************************
function WrappedText _
	(byVal vsInput, byVal viWidth)
	'vsInput             text to wrap
	'viWidth             line width

	dim   liInputLength, lsOutput, liSOL, liEOL

	'While not at the end of the input...
	liInputLength = Len(vsInput)
	lsOutput = ""
	liSOL = 1
	do while liSOL < liInputLength
	  'If the remaining input is less than one line in width...
	  if liSOL + viWidth >= liInputLength then
	    'Append the remaining input to the output.
	    lsOutput = lsOutput & Mid(vsInput, liSOL)
	    liSOL = liInputLength + 1
	  else '(the remaining input is greater than one line in width)
	    'If there is a premature line break...
	    liEOL = InStr(liSOL, vsInput, vbNewLine)
	    if (liEOL > 0) and (liSOL + viWidth >= liEOL) then
	      'Append the short line to the output.
	      lsOutput = lsOutput & Mid(vsInput, liSOL, liEOL - liSOL + 1)
	      liSOL = liEOL + Len(vbNewLine)
	    else '(there is no premature link break)
	      'If there is a blank in the line...
	      liEOL = InStrRev(vsInput, " ", liSOL + viWidth)
	      if liEOL >= liSOL then
	        'Append the line to the output and ignore leading spaces in the next line.
	        lsOutput = lsOutput & RTrim(Mid(vsInput, liSOL, liEOL - liSOL)) & vbNewLine
	        liSOL = liEOL + 1
	        do while liSOL <= liInputLength and Mid(vsInput, liSOL, 1) = " "
	          liSOL = liSOL + 1
	        loop
	      else '(there is no blank in the line)
	        'Truncate the line and append it to the output.
	        lsOutput = lsOutput & Mid(vsInput, liSOL, viWidth - 1) & "-" & vbNewLine
	        liSOL = liSOL + viWidth - 1
	      end if
	    end if
	  end if
	loop

	WrappedText = lsOutput
end function
%>