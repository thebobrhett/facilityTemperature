'****************
'Bob Rhett - Wednesday, March 16, 2011
'  Added temperature alarm notifications
'****************

dim objShell
dim objEmail
dim objdb
dim objrs
dim objupdate
dim objElement
dim objAlert
dim objMessage
dim objNotify
dim objerr
dim strSQL
dim objTT
dim objConstants
dim CommandMode
dim Sensor
dim strTXT
dim noread
dim max_noread
dim tok(4)
dim sok(4)
dim sExist(4)
dim sID(4)
dim sCRC(4)
dim sHi(4)
dim sLo(4)
dim sTemp(4)
dim strUnits
dim rec_Time
dim strMessage
dim alert_type
dim send_to

noread = 0
max_noread = 100
alert_type = ""

'on error resume next

set objShell = Wscript.CreateObject("Wscript.Shell")

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "Temp.Monitor@dorlastan.com"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Baltimore"
objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update

'Create a database connection
set objdb = CreateObject("adodb.connection")
objdb.Open "driver={MySQL ODBC 3.51 Driver};option=16387;server=Richmond.aksa.local;user=assetmgtuser;password=asset;DATABASE=asset_management;" '##MySQL w/ MyODBC v3.51
set objrs = CreateObject("adodb.recordset")
set objupdate = CreateObject("adodb.recordset")
set objElement = CreateObject("adodb.recordset")
set objAlert = CreateObject("adodb.recordset")
set objMessage = CreateObject("adodb.recordset")
set objNotify = CreateObject("adodb.recordset")

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
    do until sok(1) = true and sok(2) = true and sok(3) = true and sok(4) = true
      read_status
      wscript.sleep 5000
      read_temp
      if noread = max_noread then
        'initialize the temperature transmitter
        objShell.LogEvent EVENT_ERROR, "Error Connecting to Temperature Transmitter " & objrs("name") & " located at " & objrs("location") & ", address " & objrs("ip_address") & ":" & objrs("tcp_port") & ". Initialization in progress."
        objupdate.open "update temp_transmitter set error=2 where id=" & objrs("id"), objdb
        objTT.Disconnect
        init objrs("ip_address"), objrs("tcp_port")
        'try again
        noread = 0
        wscript.sleep 10000
        objTT.Connect objrs("ip_address"), cint(objrs("tcp_port"))
        sok(1) = false
        sok(2) = false
        sok(3) = false
        sok(4) = false
      end if
    loop
    objupdate.open "select error from temp_transmitter where id=" & objrs("id"), objdb
    if objupdate("error") = 1 then
      objupdate.close
      objShell.LogEvent EVENT_SUCCESS, "Success Connecting to Temperature Transmitter " & objrs("name") & " located at " & objrs("location") & ", address " & objrs("ip_address") & ":" & objrs("tcp_port") & "."
      objupdate.open "update temp_transmitter set error=0 where id=" & objrs("id"), objdb
    elseif objupdate("error") = 2 then
      objupdate.close
      objShell.LogEvent EVENT_SUCCESS, "Success Connecting to Temperature Transmitter " & objrs("name") & " located at " & objrs("location") & ", address " & objrs("ip_address") & ":" & objrs("tcp_port") & ". Initialization is complete."
      objupdate.open "update temp_transmitter set error=0 where id=" & objrs("id"), objdb
    else
      objupdate.close
    end if
    objTT.Disconnect
    'Modify TimeStamp for MySQL
    rec_Time = cstr(year(now)) & "-" & cstr(month(now)) & "-" & cstr(day(now)) & " " & cstr(hour(now)) & ":" & cstr(minute(now)) & ":" & cstr(second(now))
    for Sensor = 1 to 4
      if sExist(Sensor) then
        strSQL = "insert into temp_current (rec_time, transmitter, sensor, sid, crc, temperature, high, low, units) values ('" & rec_Time & "', " & objrs("id") & ", '" & Sensor & "', '" & sID(Sensor) & "', " & sCRC(Sensor) & ", " & sTemp(Sensor) & ", " & sHi(Sensor) & ", " & sLo(Sensor) & ", '" & strUnits & "')"
        objupdate.Open strSQL, objdb

'****************
'Temperature Alarm Notifications
        'Get the id of this sensor
        strSQL = "select id, location from temp_element where transmitter='" & objrs("id") & "' and channel='" & Sensor & "'"
        objElement.open strSQL, objdb
        if not objElement.eof then
          'Get the setpoints for this sensor
          wscript.echo "Element id:" & objElement("id")
          strSQL = "select * from temp_alert where id='" & objElement("id") & "'"
          objAlert.open strSQL, objdb
          if not objAlert.eof then
            'See if the current temperature violates a set point
            wscript.echo "Temperature:" & sTemp(Sensor)
            if sTemp(Sensor) > objAlert("sp_HH") then
              alert_type = "HH"
              wscript.echo "HH Setpoint:" & objAlert("sp_HH")
            elseif sTemp(Sensor) < objalert("sp_LL") then
              alert_type = "LL"
              wscript.echo "LL Setpoint:" & objAlert("sp_LL")
            elseif sTemp(Sensor) > objalert("sp_HI") then
              alert_type = "HI"
              wscript.echo "HI Setpoint:" & objAlert("sp_HI")
            elseif sTemp(Sensor) < objalert("sp_LO") then
              alert_type = "LO"
              wscript.echo "LO Setpoint:" & objAlert("sp_LO")
            end if
            if alert_type <> "" then
              'Retrieve the message for this type of alert
              wscript.echo "Alert Type:" & alert_type
              strSQL = "select alert_text from temp_message where alert_type='sp_" & alert_type & "'"
              objMessage.open strSQL, objdb
              if not objMessage.eof then
                'See who should get this alert
                wscript.echo "Alert Text:" & objMessage("alert_text")
                strSQL = "select send_to from temp_notify where id='" & objElement("id") & "' and (type_sent<>'" & alert_type & "' or sent<>'1')"
                wscript.echo "strSQL:" & strSQL
                objNotify.open strSQL, objdb
                do until objNotify.eof
                  send_to = objNotify("send_to")
                  objEmail.to = send_to
                  objEmail.subject = "Temperature Alert"
                  objEmail.textbody = objMessage("alert_text") & ". Temperature " & sTemp(Sensor) & strUnits & " at " & objElement("location") & "."
                  objEmail.send
                  wscript.echo "Sent to:" & send_to
                  wscript.echo "Text sent:" & objMessage("alert_text") & ". Temperature " & sTemp(Sensor) & strUnits & " at " & objElement("location") & ". http://web.dorlastan.com/facility/"
                  objNotify.movenext
                  strSQL = "update temp_notify set sent='1', type_sent='" & alert_type & "', ts_sent='" & rec_time & "' where send_to='" & send_to & "' and id='" & objElement("id") & "'"
                  wscript.echo "strSQL:" & strSQL
                  objupdate.open strSQL, objdb
                  strSQL = "update temp_alert set ts_" & alert_type & "='" & rec_time & "'"
                  wscript.echo "strSQL:" & strSQL
                  objupdate.open strSQL, objdb
                loop
                objNotify.close
              end if
              objMessage.close
            end if
          end if
          objAlert.close
        end if
        objElement.close
'****************
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
set objShell=nothing

'****************

sub read_status()
  wscript.echo "Reading Status"
  objTT.SendString "?"
  strTXT = objTT.ReceiveString
  wscript.echo "Status Received:" & strTXT
  strStatus = split(strTXT, ",")
  if ubound(strStatus) > 0 then
    strUnits = trim(strStatus(3))
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
end sub

sub parse_status(Sensor, i, strStatus)
  wscript.echo "Parsing Status:" & strStatus(i)
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

sub read_temp
  wscript.echo "Reading Temperature"
  objTT.SendString "!"
  do until sok(1) = true and sok(2) = true and sok(3) = true and sok(4) = true
    strTXT = objTT.ReceiveString
    wscript.echo "Temperature Received:" & strTXT
    strFeed = split(strTXT, ",")
    if ubound(strFeed) > 0 then
      for j = 0 to ubound(strFeed)
        select case right(strFeed(j), 8)
          case "Sensor 1"
            if not sok(1) then
              wscript.echo "Sensor 1"
              parse_temp 1, j, strFeed
            end if
          case "Sensor 2"
            if not sok(2) then
              wscript.echo "Sensor 2"
              parse_temp 2, j, strFeed
            end if
          case "Sensor 3"
            if not sok(3) then
              wscript.echo "Sensor 3"
              parse_temp 3, j, strFeed
            end if
          case "Sensor 4"
            if not sok(4) then
              wscript.echo "Sensor 4"
              parse_temp 4, j, strFeed
            end if
        end select
      next
    end if
    noread = noread + 1
    if noread = max_noread then exit do
  loop
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

sub init(ip_address, tcp_port)
  dim objTS
  dim console_port
  dim ts_port

  console_port = 23
  ts_port = right(cstr(tcp_port), 2)

  'Create a socket instance
  set objTS = CreateObject("ActiveXperts.Tcp")
  objTS.Protocol = objConstants.asSOCKET_PROTOCOL_RAW

  'Connect to the ports
  wscript.echo "Connecting to " & ip_address & ":" & cstr(console_port)
  objTS.Connect ip_address, console_port

  if objTS.lasterror <> 0 then
    wscript.echo "Error Connecting to Terminal Server console port."
  else
    wscript.echo "Connected to Terminal Server console port."
    objTS.SendString vbCrLf
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "ACCESS"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "BATCH"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "SET PRIV"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "SYSTEM"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "DEFINE PORT " & ts_port & " MODEM ENABLED DTRWAIT ENABLED"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "LOGOUT PORT " & ts_port
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    wscript.sleep 10000

    objTS.SendString "DEFINE PORT " & ts_port & " MODEM DISABLED DTRWAIT DISABLED"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "LOGOUT PORT " & ts_port
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop

    objTS.SendString "LOGOUT"
    do while objTS.HasData
      strReceived = objTS.ReceiveString
      WScript.Echo strReceived
    loop
  end if

  objTS.disconnect

  objTT.Connect objrs("ip_address"), cint(objrs("tcp_port"))
  do while objTT.HasData
    strReceived = objTT.ReceiveString
    WScript.Echo strReceived
  loop
  objTT.disconnect

  set objTS=nothing
end sub
