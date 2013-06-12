dim objConn
dim strSQL
dim objCom
dim waiting
dim strTXT

'on error resume next

set objCom = WScript.CreateObject("MSWinsock.Winsock", "WSEvent_")  

objCom.Connect "10.75.10.229", 23
waiting=true
do while waiting
  wscript.sleep 1000
loop

objCom.Close

Sub WSEvent_Connect()
  objCom.SendData vbCrLf
  objCom.SendData "access"
  objCom.SendData "batch"
  objCom.SendData "set priv"
  objCom.SendData "system"
  objCom.SendData "def port 4 dtrwait ena"
  objCom.SendData "lo port 4"
  objCom.SendData "lo"
  waiting=false
end sub

Sub WSEvent_DataArrival(bytes)
  strResponse ="empty"
  do until strResponse = ""
    objCom.GetData strResponse
    wscript.echo "Response: " & strResponse
  loop
end sub
