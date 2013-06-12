dim objConn
dim objRStt
dim objRS
dim strSQL
dim objCom
dim strTT
dim rec_Time
dim keepSleeping
dim CommandMode
'dim strTXT           'make these local
'dim strStatus        'make these local
'dim strFeed          'make these local
'dim i                'make these local
dim strUnits
dim sID(4)
dim sCRC(4)
dim sHi(4)
dim sLo(4)
dim sTemp(4)
dim sok(4)
dim sExist(4)

on error resume next

MSG_TO = "bob.rhett@dorlastan.com, john.stasiek@dorlastan.com, keith.brooks@dorlastan.com"
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "Asset.Management@dorlastan.com"
objEmail.To = MSG_TO
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Baltimore"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update

set objConn = CreateObject("adodb.connection")
objConn.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objRS = CreateObject("adodb.recordset")
set objRStt = CreateObject("adodb.recordset")

set objCom = WScript.CreateObject("MSWinsock.Winsock", "WSEvent_")  

strSQL = "select * from temp_transmitter"
objRStt.open strSQL, objConn

do until objRStt.eof
  strTT = objRStt("name")
  wscript.echo "Transmitter: " & strTT
  rec_Time = Now()
  wscript.echo "TimeStamp: " & rec_Time
  objCom.Connect objRStt("ip_address"), objRStt("tcp_port")

  keepSleeping=true
  while keepSleeping
    WScript.Sleep 1000
'Check for connection
  wend

  objCom.SendData "!"
  CommandMode = false
  wscript.echo
  wscript.echo "Current Temperature"

  do until sok(1) = true and sok(2) = true and sok(3) = true and sok(4) = true
    WScript.Sleep 2000
  loop

  objCom.Close

  for Sensor = 1 to 4
    if sExist(Sensor) then
      'Modify TimeStamp for MySQL
      rec_Time = cstr(year(rec_Time)) & "-" & cstr(month(rec_Time)) & "-" & cstr(day(rec_Time)) & " " & cstr(hour(rec_Time)) & ":" & cstr(minute(rec_Time)) & ":" & cstr(second(rec_Time))
      strSQL = "insert into temp_current (rec_time, transmitter, sensor, sid, crc, temperature, high, low, units) values ('" & rec_Time & "', " & objRStt("id") & ", '" & Sensor & "', '" & sID(Sensor) & "', " & sCRC(Sensor) & ", " & sTemp(Sensor) & ", " & sHi(Sensor) & ", " & sLo(Sensor) & ", '" & strUnits & "')"
      wscript.echo strSQL
      objRS.Open strSQL, objConn
    end if
  next
  objRStt.movenext
loop

objConn.close
set objRStt = nothing
set objRS = nothing
set objConn = nothing
wscript.echo "****************"

'****************

Sub WSEvent_Connect()
'Port opened
  objCom.SendData "X"
  objCom.SendData "?"
  CommandMode = true
end sub

Sub WSEvent_DataArrival(bytes)
  dim strTXT1
  dim strTXT2
  dim strStatus
  dim strFeed
  dim i
  dim j
  if CommandMode then
    strTXT1 = ""
    objCom.GetData strTXT1
    strStatus = split(strTXT1, ",")
    if ubound(strStatus) > 0 then
      strUnits = trim(strStatus(ubound(strStatus) - 1))
      for i = 0 to ubound(strStatus)
        select case right(strStatus(i), 8)
          case "Sensor 1"
            wscript.echo "Sensor 1"
            parse_status 1, i, strStatus
          case "Sensor 2"
             wscript.echo "Sensor 2"
            parse_status 2, i, strStatus
          case "Sensor 3"
            wscript.echo "Sensor 3"
            parse_status 3, i, strStatus
          case "Sensor 4"
            wscript.echo "Sensor 4"
            parse_status 4, i, strStatus
        end select
      next
    end if
  else
    strTXT2 = ""
    objCom.GetData strTXT2
    strFeed = split(strTXT2, ",")
    if ubound(strFeed) > 0 then
      for j = 0 to ubound(strFeed)
        select case right(strFeed(j), 7)
          case "ensor 1"
            if not sok(1) then
              wscript.echo "Sensor 1"
              parse_temp 1, j, strFeed
            end if
          case "ensor 2"
            if not sok(2) then
              wscript.echo "Sensor 2"
              parse_temp 2, j, strFeed
            end if
          case "ensor 3"
            if not sok(3) then
              wscript.echo "Sensor 3"
              parse_temp 3, j, strFeed
            end if
          case "ensor 4"
            if not sok(4) then
              wscript.echo "Sensor 4"
              parse_temp 4, j, strFeed
            end if
        end select
      next
    end if
  end if
  keepSleeping = false
end sub

sub parse_status(Sensor, i, strStatus)
  if left(trim(strStatus(i + 1)), 4) <> "Open" then
    if left(trim(strStatus(i + 1)), 5) <> "Short" then
      if mid(strStatus(i + 1), 18, 19) <> "00-OK-00 CRC Errors" then
        sID(Sensor) = left(trim(strStatus(i + 1)), 16)
        sCRC(Sensor) = true
        sHi(Sensor) = csng(trim(strStatus(i + 2)))
        sLo(Sensor) = csng(trim(strStatus(i + 5)))
        wscript.echo "ID: " & sID(Sensor)
        wscript.echo "CRC: ok"
        wscript.echo "Hi: " & sHi(Sensor)
        wscript.echo "Lo: " & sLo(Sensor)
      else
        wscript.echo "Bad"
      end if
    else
      wscript.echo "Short"
    end if
  else
    sok(Sensor) = true
    wscript.echo "Open"
  end if
end sub

sub parse_temp(Sensor, j, strFeed)
  sTemp(Sensor) = csng(trim(strFeed(j + 1)))
  if sHi(Sensor) = 185 then
    sHi(Sensor) = sTemp(Sensor)     'Ignore Hi Temp of 185 (reconnection of a sensor)
  end if
  sok(Sensor) = true
  sExist(Sensor) = true
  wscript.echo "Temp: " & sTemp(Sensor) & " " & strUnits
end sub