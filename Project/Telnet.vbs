' ********************************************************************
' ActiveSocket sample - Telnet client communication
'    Connects to www.activexperts.com and reads page /activsocket/demopage
' (c) Copyright 2006 by ActiveXperts Software 
'    http://www.activexperts.com
' ********************************************************************

Option Explicit

' Declare variables
Dim objTcp, objConst, strReceived, ip_address, tcp_port

ip_address = "10.75.10.229"
tcp_port = 23

' Create a Tcp instance
Set objTcp  	= CreateObject ( "ActiveXperts.Tcp" )
Set objConst	= CreateObject ( "ActiveXperts.ASConstants" )

objTcp.Protocol = objConst.asSOCKET_PROTOCOL_TELNET

' Write some information to console
WScript.Echo "ActiveSocket " & objTcp.Version & " demo."
WScript.Echo "Expiration date: " & objTcp.ExpirationDate & vbCrLf

' Make a connection on port 80 to remote server
'objTcp.Connect "www.activexperts.com", 80
objTcp.Connect ip_address, tcp_port
If objTcp.LastError <> 0 Or objTcp.ConnectionState <> objConst.asSOCKET_CONNSTATE_CONNECTED Then
  Wscript.Echo "Error connecting to " & ip_address & ":" & tcp_port & ", result: " & objTcp.LastError
  WScript.Quit
End If

' YES, connection established.
Wscript.Echo "Connected to " & ip_address & ":" & tcp_port & ", result: " & objTcp.LastError
	
objTcp.SendString vbCrLf
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "access"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "batch"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "set priv"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "system"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "def port 4 autoprompt ena"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "lo port 4"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "sho port 4"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop

objTcp.SendString "lo"
do while objTcp.HasData
  strReceived = objTcp.ReceiveString
  WScript.Echo strReceived
loop
	
objTcp.Disconnect
WScript.Echo "Disconnected"
