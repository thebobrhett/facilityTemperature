'****************
'Bob Rhett - Thursday, March 17, 2011
'****************

dim objShell
dim objEmail
dim objdb
dim objElement
dim objAlert
dim objNotify
dim strSQL

'on error resume next

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "Temp.Monitor@dorlastan.com"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Baltimore"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update

'Create a database connection
set objdb = CreateObject("adodb.connection")
objdb.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objElement = CreateObject("adodb.recordset")
set objAlert = CreateObject("adodb.recordset")
set objNotify = CreateObject("adodb.recordset")

'Temperature Notification Test
'Get list of subscribers and elements to which they are subscribed
strSQL = "select send_to, id from temp_notify"
objNotify.open strSQL, objdb
do until objNotify.eof
  'Get the transmitter and sensor number for the element
  strSQL = "select transmitter, channel, location from temp_element where id='" & objNotify("id") & "'"
  objElement.open strSQL, objdb
  if not objElement.eof then
    'Get the last recorded temperature and time for the element
    strSQL = "select temperature, rec_time, units from temp_current where transmitter='" & objElement("transmitter") & "' and sensor='" & objElement("channel") & "' order by rec_time desc"
    objAlert.open strSQL, objdb
    if not objAlert.eof then
      objEmail.to = objNotify("send_to")
      objEmail.subject = "Temperature Alert Test"
      objEmail.textbody = "This is a test. Temperature " & objAlert("temperature") & objAlert("units") & " at " & objElement("location") & " recorded at " & formatdatetime(objAlert("rec_time")) & ". http://web.dorlastan.com/facility/"
      objEmail.send
    end if
    objAlert.close
  end if
  objElement.close
  objNotify.movenext
loop
objNotify.close

objdb.close
set objElement=nothing
set objAlert=nothing
set objNotify=nothing
set objdb=nothing
set objEmail=nothing
