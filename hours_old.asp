<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
</meta><title>Facility Temperatures</title>
<link rel=STYLESHEET href='http://mogsa4/aksastyle.css' type='text/css'>
<style type='text/css'>
<!--
a:link     { color:black; text-decoration:underline; }
a:visited  { color:black; text-decoration:underline; }
a img { border-width: 1; }
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
Const pasthrs = -24

dim objConn
dim objTransmitter
dim objElement
dim objHourly
dim objCurrent
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
set objCurrent = CreateObject("adodb.recordset")

'determine start and end times
if request("d") <> "" then
  strStart = cstr(year(dateadd("h", pasthrs, request("d")))) & "-" & cstr(month(dateadd("h", pasthrs, request("d")))) & "-" & cstr(day(dateadd("h", pasthrs, request("d")))) & " " & cstr(hour(dateadd("h", pasthrs, request("d")))) & ":00"
  strEnd = cstr(year(request("d"))) & "-" & cstr(month(request("d"))) & "-" & cstr(day(request("d"))) & " " & cstr(hour(request("d"))) & ":00"
else
  strStart = cstr(year(dateadd("h", pasthrs, now))) & "-" & cstr(month(dateadd("h", pasthrs, now))) & "-" & cstr(day(dateadd("h", pasthrs, now))) & " " & cstr(hour(dateadd("h", pasthrs, now))) & ":00"
  strEnd = cstr(year(now)) & "-" & cstr(month(now)) & "-" & cstr(day(now)) & " " & cstr(hour(now)) & ":00"
end if
intLoScale = 60
intHiScale = 80
barwidth = 80/pasthrs

response.write "<h1/><p class='center'>Facility Temperature Trends (past " & 0 - pasthrs & " hours)"
strSQL = "select * from temp_transmitter"
objTransmitter.open strSQL, objConn

do until objTransmitter.eof
  response.write "<table width='80%' cellpadding='0' cellspacing='0' border='1'>"
  response.write "<tr><td bgcolor='powderblue'><h2/>Location: " & objTransmitter("location") & "</td></tr><tr><td>"
  strSQL = "select * from temp_element where transmitter = " & objTransmitter("id") & " order by transmitter, channel"
  objElement.open strSQL, objConn
  do until objElement.eof
    response.write "<table width='100%' cellpadding='0' cellspacing='0'>"
    strSQL = "select rec_time, temperature from temp_current where transmitter = " & objTransmitter("id") & " and sensor = " & objElement("channel") & " order by rec_time desc"
    objCurrent.open strSQL, objConn
    response.write "<tr><td rowspan='2' align='right' valign='top'>"

    response.write "<table width='100%' cellpadding='0' cellspacing='0'>"
    response.write "<tr><td valign='top' align='right'>"
    response.write "<h3/>" & intHiScale & " DegF<br/>&nbsp;</td></tr>"
    response.write "<tr><td valign='bottom' align='center'>"
    response.write "<h3/>&nbsp;<br/>Temperature last read at<br/>" & objCurrent("rec_time") & "<br/><b>" & objCurrent("temperature") & " DegF</b></td></tr>"
    response.write "</table>"


    response.write "</td></tr>"



    objCurrent.close
    strSQL = "select * from temp_hourly where transmitter = " & objTransmitter("id") & " and sensor = " & objElement("channel") & " and rec_time between '" & strStart & "' and '" & strEnd & "' order by rec_time"
    objHourly.open strSQL, objConn
    response.write "<tr>"
    do until objHourly.eof

      if objHourly("temp_lo") => intHiScale then
        intMIN = intHiScale - intLoScale
      elseif objHourly("temp_lo") => intLoScale then
        intMIN = objHourly("temp_lo") - intLoScale
      else
        intMIN = 0
      end if

      if intMIN < 0 then intMIN = 0

      if objHourly("temp_avg") => intHiScale then
        intAVG = intHiScale - intLoScale
      elseif objHourly("temp_avg") => intLoScale then
        intAVG = objHourly("temp_avg") - intLoScale - intMIN
      else
        intAVG = 0
      end if

      if intAVG < 0 then intAVG = 0

      if objHourly("temp_hi") => intHiScale then
        intMAX = intHiScale - intLoScale
      elseif objHourly("temp_hi") => intLoScale then
        intMAX = objHourly("temp_hi") - intLoScale - intMIN - intAVG
      else
        intMAX = 0
      end if

      if intMAX < 0 then intMAX = 0

      response.write "<td valign='bottom'>"
      response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/trans.gif' width='100%' height='" & (intHiScale - intMAX - intAVG - intMIN) * 2 & "' border='0'><br/>"
      response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/red.gif' alt='High Temp " & objHourly("temp_hi") & " DegF @" & FormatDateTime(objHourly("rec_time")) & "' width='100%' height='" & intMAX * 2 & "' border='1'><br/>"
      response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/green.gif' alt='Average Temp " & FormatNumber(objHourly("temp_avg"), 2) & " DegF @" & FormatDateTime(objHourly("rec_time")) & "' width='100%' height='" & intAVG * 2 & "' border='1'><br/>"
      response.write "<img src='http://mogsb8/assetmgt/VerticalGraph/blue.gif' alt='Low Temp " & objHourly("temp_lo") & " DegF @" & FormatDateTime(objHourly("rec_time")) & "' width='100%' height='" & intMIN * 2 & "' border='1'>"
      response.write "</td>"

      objHourly.movenext
    loop
    response.write "</tr>"
    response.write "<tr><td align='right' valign='bottom'><h3/>" & intLoScale & " DegF</td><th colspan='" & 0-pasthrs & "' bgcolor='powderblue'><h2/>" & objElement("location") & "</th></tr>"
    response.write "<tr><td colspan='" & 1-pasthrs & "'>&nbsp;<hr/></td></tr></table>"
    objHourly.close
    objElement.movenext
  loop
  objElement.close
  response.write "</td></tr></table>"
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