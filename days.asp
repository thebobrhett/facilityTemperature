<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
</meta><title>Facility Temperatures</title>
<link rel=STYLESHEET href='http://mogsa4/aksastyle.css' type='text/css'>
<style type='text/css'>
<!--
a:link     { color:black; text-decoration:underline; }
a:visited  { color:black; text-decoration:underline; }
-->
</style>
</head>
<body>

<%
'****************
'Bob Rhett - Thursday, July 24, 2008
'  Created
'****************
'on error resume next
Const pastdays = -30

dim objConn
dim objTransmitter
dim objElement
dim objHourly
dim strStart
dim strEnd
dim strSQL
dim intMIN
dim intAVG
dim intMAX
dim inLoScale
dim intHiScale

set objConn = CreateObject("adodb.connection")
objConn.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objTransmitter = CreateObject("adodb.recordset")
set objElement = CreateObject("adodb.recordset")
set objHourly = CreateObject("adodb.recordset")

'determine start and end times
intLoScale = 60
intHiScale = 80
barheight = 4

response.write "<h1/><p class='center'>Facility Temperature Trends (past " & 0 - pastdays & " days)"
strSQL = "select * from temp_transmitter"
objTransmitter.open strSQL, objConn

do until objTransmitter.eof
  response.write "<table width='80%' cellpadding='0' cellspacing='0'>"
  response.write "<tr><td bgcolor='powderblue'><h2/>" & objTransmitter("location") & "</td></tr><tr><td>"
  strSQL = "select * from temp_element where transmitter = " & objTransmitter("id") & " order by transmitter, channel"
  objElement.open strSQL, objConn
  do until objElement.eof
    response.write "<table width='100%' cellpadding='0' cellspacing='0'><tr><td rowspan='2' align='right' valign='top'><h3/>" & intHiScale & " DegF</td></tr>"
    response.write "<tr>"
    for i = pastdays to -1
      strStart = cstr(year(dateadd("d", i, now))) & "-" & cstr(month(dateadd("d", i, now))) & "-" & cstr(day(dateadd("d", i, now)))
      strEnd = cstr(year(dateadd("d", i + 1, now))) & "-" & cstr(month(dateadd("d", i + 1, now))) & "-" & cstr(day(dateadd("d", i + 1, now)))
      strSQL = "select units, avg(temp_avg) as AVG, max(temp_hi) as MAX, min(temp_lo) as MIN from temp_hourly where transmitter = " & objTransmitter("id") & " and sensor = " & objElement("channel") & " and rec_time between '" & strStart & "' and '" & strEnd & "' group by units"
      objHourly.open strSQL, objConn
      do until objHourly.eof

        if objHourly("MIN") > intHiScale then
          intMIN = intHiScale - intLoScale
        elseif objHourly("MIN") > intLoScale then
          intMIN = objHourly("MIN") - intLoScale
        else
          intMIN = 0
        end if

        if intMIN < 0 then intMIN = 0

        if objHourly("AVG") > intHiScale then
          intAVG = intHiScale - intLoScale - intMIN
        elseif objHourly("AVG") > intLoScale then
          intAVG = objHourly("AVG") - intLoScale - intMIN
        else
          intAVG = 0
        end if

        if intAVG < 0 then intAVG = 0

        if objHourly("MAX") > intHiScale then
          intMAX = intHiScale - intLoScale - intMIN - intAVG
        elseif objHourly("MAX") > intLoScale then
          intMAX = objHourly("MAX") - intLoScale - intMIN - intAVG
        else
          intMAX = 0
        end if

        if intMAX < 0 then intMAX = 0

        response.write "<td valign='bottom'>"
        response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/trans.gif' width='100%' height='" & (intHiScale - intLoScale - intMAX - intAVG - intMIN) * barheight & "' border='0'><br/>"
        response.write "<a href='http://mogsb8/assetmgt/temperature/hours.asp?d=" & FormatDateTime(strStart) & "'>"
        response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/red.gif' alt='High Temp " & objHourly("MAX") & " DegF @" & FormatDateTime(strStart) & "' width='100%' height='" & intMAX * barheight & "' border='1'><br/>"
        response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/green.gif' alt='Average Temp " & FormatNumber(objHourly("AVG"), 2) & " DegF @" & FormatDateTime(strStart) & "' width='100%' height='" & intAVG * barheight & "' border='1'><br/>"
        response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/blue.gif' alt='Low Temp " & objHourly("MIN") & " DegF @" & FormatDateTime(strStart) & "' width='100%' height='" & intMIN * barheight & "' border='1'>"
        response.write "</a>"
        response.write "</td>"

        objHourly.movenext
      loop
      objHourly.close
    next
    response.write "</tr>"
    response.write "<tr><td align='right' valign='bottom'><h3/>" & intLoScale & " DegF</td><th colspan='" & 0-pastdays & "' bgcolor='powderblue'><h2/>" & objElement("location") & "</th></tr>"
    response.write "<tr><td colspan='" & 1-pastdays & "'>&nbsp;<hr/></td></tr>"
    objElement.movenext
  loop
  objElement.close
  response.write "</table>"
  objTransmitter.movenext
loop

objTransmitter.close
objConn.close

set objHourly = nothing
set objElement = nothing
set objTransmitter = nothing
set objCurrent = nothing
set objConn = nothing
%>

</p>
</body>
</html>