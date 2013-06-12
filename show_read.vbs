dim objConn
dim objRStt
dim objRS
dim strSQL
dim objCom
dim strTT
dim rec_Time
dim waiting
dim CommandMode
dim strTXT

'on error resume next

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
  waiting=true

  objCom.Connect objRStt("ip_address"), objRStt("tcp_port")

  while waiting
    WScript.Sleep 1000
  wend

  objCom.Close

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
  WSCRIPT.ECHO "Port Opened"
  objCom.SendData "?"
end sub

Sub WSEvent_DataArrival(bytes)
  strTXT="empty"
  objCom.GetData strTXT
  wscript.echo "*START*"
  wscript.echo strTXT
  wscript.echo "*END*"
  waiting = false
end sub

Sub WSEvent_Close()
end sub
