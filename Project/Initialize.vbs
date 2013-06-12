Set objComport = CreateObject( "ActiveXperts.Comport" )

objComport.Device     = "COM1"
objComport.Baudrate   = 9600
objComport.DataBits   = 8
objComport.Stopbits   = 1
objComport.Parity     =  1   'None
objComport.ComTimeout = 500
objComport.HardwareFlowControl = 1   'None
objComport.SoftwareFlowControl = 1   'None

objComport.Open ()

if objComport.LastError = 0 and objComport.IsOpened = -1 Then
  objComport.ComTimeout = 10000  'Function will timeout after 10sec
  str = objComport.ReadString
  if str="" then
    wscript.echo "No Readings"
  else
    objComport.ComTimeout = 1000  'Function will timeout after 1sec
    wscript.echo " <-?"
    objComport.WriteString "?"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-D"
    objComport.WriteString "D"
    str = objComport.ReadString
    wscript.echo " ->" & str
    wscript.echo " <-?"
    objComport.WriteString "?"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-S"
    objComport.WriteString "S"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-N"
    objComport.WriteString "N"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-S"
    objComport.WriteString "S"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-N"
    objComport.WriteString "N"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-N"
    objComport.WriteString "N"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-N"
    objComport.WriteString "N"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-N"
    objComport.WriteString "N"
    str = "notempty"
    do while str <> ""
      str = objComport.ReadString
      wscript.echo " ->" & str
    loop
    wscript.echo " <-X"
    objComport.WriteString "X"
  end if
else
  wScript.echo "IsOpened = " & objComport.IsOpened
  wscript.echo "Error = " & objComport.LastError
  wScript.echo "Error description: " & objComport.GetErrorDescription( objComport.LastError )
end If

objComport.Close ()
