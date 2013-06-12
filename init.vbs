dim objTS
dim objConstants
dim ip_address
dim console_port
dim tcp_port

ip_address = "10.75.10.229"
console_port = 23
tcp_port = 2002

'on error resume next

'Create a socket instance
set objTS = CreateObject("ActiveXperts.Tcp")
set objConstants = CreateObject("ActiveXperts.ASConstants")
objTS.Protocol = objConstants.asSOCKET_PROTOCOL_RAW

'Connect to the ports
wscript.echo "Connecting to " & ip_address & ":" & cstr(console)
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

  objTS.SendString "DEFINE PORT 2 MODEM ENABLED DTRWAIT ENABLED"
  do while objTS.HasData
    strReceived = objTS.ReceiveString
    WScript.Echo strReceived
  loop

  objTS.SendString "LOGOUT PORT 2"
  do while objTS.HasData
    strReceived = objTS.ReceiveString
    WScript.Echo strReceived
  loop

  wscript.sleep 10000

  objTS.SendString "DEFINE PORT 2 MODEM DISABLED DTRWAIT DISABLED"
  do while objTS.HasData
    strReceived = objTS.ReceiveString
    WScript.Echo strReceived
  loop

  objTS.SendString "LOGOUT PORT 2"
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
wscript.echo "Connecting to " & ip_address & ":" & cstr(tcp_port)
objTS.Connect ip_address, tcp_port

if objTS.lasterror <> 0 then
  wscript.echo "Error Connecting to Temperature Transmitter."
else
  wscript.echo "Connected to Temperature Transmitter."
  objTS.SendString "LOGOUT"
  do while objTS.HasData
    strReceived = objTS.ReceiveString
    WScript.Echo strReceived
  loop
end if

objTS.disconnect

set objConstants=nothing
set objTS=nothing
