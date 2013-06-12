dim objShell
dim objEmail
dim msgTo
dim objdb
dim objrs
dim objerr
dim strSQL
dim objTT
dim objConstants
dim CommandMode
dim Sensor
dim sok(4)
dim sExist(4)
dim sID(4)
dim sCRC(4)
dim sHi(4)
dim sLo(4)
dim sTemp(4)
dim strUnits
dim rec_Time

dim done
dim strTXT
dim strMessage

'on error resume next

set objShell = Wscript.CreateObject("Wscript.Shell")

'msgTo = "bob.rhett@dorlastan.com, john.stasiek@dorlastan.com, keith.brooks@dorlastan.com"
msgTo = "bob.rhett@dorlastan.com"
Set objEmail = CreateObject("CDO.Message")
objEmail.From = "Asset.Management@dorlastan.com"
objEmail.To = msgTo
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Baltimore"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update

'Create a databse connection
set objdb = CreateObject("adodb.connection")
objdb.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objrs = CreateObject("adodb.recordset")
set objupdate = CreateObject("adodb.recordset")

'Create a socket instance
set objTT = CreateObject("ActiveXperts.Tcp")
set objConstants = CreateObject("ActiveXperts.ASConstants")
objTT.Protocol = objConstants.asSOCKET_PROTOCOL_RAW

strSQL = "select * from temp_transmitter where ip_address is not null"
objrs.open strSQL, objdb

do until objrs.eof
  'Connect to the ports
  wscript.echo "Connecting to " & objrs("name") & " at " & objrs("ip_address") & ":" & objrs("tcp_port")
  objTT.Connect objrs("ip_address"), cint(objrs("tcp_port"))
  if objTT.lasterror <> 0 then
    objShell.LogEvent EVENT_ERROR, "Error Connecting to Temperature Transmitter " & objrs("name") & " located at " & objrs("location") & ", address " & objrs("ip_address") & ":" & objrs("tcp_port")
    objupdate.open "update temp_transmitter set error=1 where id=" & objrs("id"), objdb
  else
    wscript.echo "Connected"
    objTT.SendString "?"
    strTXT1 = objTT.ReceiveString
    strStatus = split(strTXT1, ",")
    if ubound(strStatus) > 0 then
      strUnits = trim(strStatus(ubound(strStatus) - 1))
      for i = 0 to ubound(strStatus)
        select case right(strStatus(i), 7)
          case "ensor 1"
            wscript.echo "Sensor 1"
            parse_status 1, i, strStatus
          case "ensor 2"
            wscript.echo "Sensor 2"
            parse_status 2, i, strStatus
          case "ensor 3"
            wscript.echo "Sensor 3"
            parse_status 3, i, strStatus
          case "ensor 4"
            wscript.echo "Sensor 4"
            parse_status 4, i, strStatus
        end select
      next
    end if
    objTT.SendString "!"
    do until sok(1) = true and sok(2) = true and sok(3) = true and sok(4) = true
      strTXT2 = objTT.ReceiveString
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
    loop
    objTT.Disconnect
    for Sensor = 1 to 4
      if sExist(Sensor) then
        'Modify TimeStamp for MySQL
        rec_Time = cstr(year(now)) & "-" & cstr(month(now)) & "-" & cstr(day(now)) & " " & cstr(hour(now)) & ":" & cstr(minute(now)) & ":" & cstr(second(now))
        strSQL = "insert into temp_current (rec_time, transmitter, sensor, sid, crc, temperature, high, low, units) values ('" & rec_Time & "', " & objrs("id") & ", '" & Sensor & "', '" & sID(Sensor) & "', " & sCRC(Sensor) & ", " & sTemp(Sensor) & ", " & sHi(Sensor) & ", " & sLo(Sensor) & ", '" & strUnits & "')"
        objupdate.Open strSQL, objdb
      end if
    next
  end if
  objrs.movenext
loop

objrs.close
objdb.close
set objupdate=nothing
set objrs=nothing
set objdb=nothing
set objConstants=nothing
set objTT=nothing
set objEmail=nothing
set obhShell=nothing

'****************

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
      strUnits = trim(strStatus(ubound(strStatus) - 2))
      for i = 0 to ubound(strStatus)
        select case right(strStatus(i), 7)
          case "ensor 1"
            wscript.echo "Sensor 1"
            parse_status 1, i, strStatus
          case "ensor 2"
             wscript.echo "Sensor 2"
            parse_status 2, i, strStatus
          case "ensor 3"
            wscript.echo "Sensor 3"
            parse_status 3, i, strStatus
          case "ensor 4"
            wscript.echo "Sensor 4"
            parse_status 4, i, strStatus
        end select
      next
    end if
  else
    strTXT2 = ""
    objTT.GetData strTXT2
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