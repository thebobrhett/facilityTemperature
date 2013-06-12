'****************
'Bob Rhett - Friday, December 11, 2009
'  Changed method for claculating the start time. Previous method missed the hour before midnight.
'****************

dim strStart
dim strEnd
dim Keep
dim strDelEnd
dim objConn
dim objCurrent
dim objTransmitter
dim objElement
dim objHourly
dim strSQL

on error resume next

set objConn = CreateObject("adodb.connection")
objConn.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objCurrent = CreateObject("adodb.recordset")
set objTransmitter = CreateObject("adodb.recordset")
set objElement = CreateObject("adodb.recordset")
set objHourly = CreateObject("adodb.recordset")

Keep = -30

'determine start and end times
strStart = dateadd("h", -1, now)
strStart = cstr(year(strStart)) & "-" & cstr(month(strStart)) & "-" & cstr(day(strStart)) & " " & cstr(hour(strStart)) & ":00"
strEnd = cstr(year(now)) & "-" & cstr(month(now)) & "-" & cstr(day(now)) & " " & cstr(hour(now)) & ":00"
strDelEnd = cstr(year(dateadd("d", Keep, now))) & "-" & cstr(month(dateadd("d", Keep, now))) & "-" & cstr(day(dateadd("d", Keep, now)))

strSQL = "select * from temp_transmitter"
objTransmitter.open strSQL, objConn

do until objTransmitter.eof
  strSQL = "select * from temp_element where transmitter = " & objTransmitter("id") & " order by channel"
  objElement.open strSQL, objConn
  do until objElement.eof
    strSQL = "select units, avg(temperature) as AVG, max(high) as MAX, min(low) as MIN from temp_current where transmitter = " & objTransmitter("id") & " and sensor = " & objElement("channel") & " and crc = 1 and rec_time between '" & strStart & "' and '" & strEnd & "' group by units"
    objCurrent.open strSQL, objConn
    strSQL = "insert into temp_hourly (rec_time, transmitter, sensor, temp_avg, temp_hi, temp_lo, units) values ('" & strEnd & "', " & objTransmitter("id") & ", " & objElement("channel") & ", " & FormatNumber(objCurrent("AVG"), 2) & ", " & objCurrent("MAX") & ", " & objCurrent("MIN") & ", '" & objCurrent("units") & "')"
    objHourly.open strSQL, objConn
    objCurrent.close
    strSQL = "delete from temp_current where transmitter = " & objTransmitter("id") & " and sensor = " & objElement("channel") & " and rec_time < '" & strDelEnd & "'"
    objCurrent.open strSQL, objConn
    objElement.movenext
  loop
  objElement.close
  objTransmitter.movenext
loop

objTransmitter.close
objConn.close

set objHourly = nothing
set objElement = nothing
set objTransmitter = nothing
set objCurrent = nothing
set objConn = nothing
