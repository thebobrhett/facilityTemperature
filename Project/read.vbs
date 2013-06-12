on error resume next

dim Field, Sensor1, Sensor2, Sensor3, Sensor4

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
  objComport.ComTimeout = 2000  'Function will timeout after 2sec
  str = objComport.ReadString
  if str="" then
    wscript.echo "No Readings"
  else
'    wScript.echo str
    Field = split(str,",")
    objComport.WriteString "?"
    str = objComport.ReadString
    str = objComport.ReadString
    Sensor1 = split(str,",")
    str = objComport.ReadString
    Sensor2 = split(str,",")
    str = objComport.ReadString
    Sensor3 = split(str,",")
    str = objComport.ReadString
    Sensor4 = split(str,",")
    do while str <> ""
      str = objComport.ReadString
    loop
    objComport.WriteString "!"

    wscript.echo "Device: " & objComport.Device
'Add Timestamp
'Ignore Hi Temp of 185 (reconnection or a sensor)
    if trim(Sensor1(1)) <> "Open" and trim(Sensor1(1)) <> "Short" then
      wscript.echo "Sensor: " & Field(0)
      wscript.echo "s/n: " & left(trim(Sensor1(1)), 16)
      wscript.echo "Current Temp: " & trim(Field(1))
      wscript.echo "Hi Temp: " & trim(Field(4))
      wscript.echo "Lo Temp: " & trim(Field(7))
      wscript.echo "Units: " & trim(Sensor1(3))
      if mid(Sensor1(1), 18, 19) <> "00-OK-00 CRC Errors" then
        wscript.echo "Data Integrity: OK"
      else
        wscript.echo "Data Integrity: Bad"
      end if
    elseif trim(Sensor1(1)) = "Short" then
      wscript.echo "Sensor 1: Short"
    end if
    if trim(Sensor2(1)) <> "Open" and trim(Sensor2(1)) <> "Short" then
      wscript.echo "Sensor: " & Field(0)
      wscript.echo "s/n: " & left(trim(Sensor2(1)), 16)
      wscript.echo "Current Temp: " & trim(Field(1))
      wscript.echo "Hi Temp: " & trim(Field(4))
      wscript.echo "Lo Temp: " & trim(Field(7))
      wscript.echo "Units: " & trim(Sensor2(3))
      if mid(Sensor2(1), 18, 19) <> "00-OK-00 CRC Errors" then
        wscript.echo "Data Integrity: OK"
      else
        wscript.echo "Data Integrity: Bad"
      end if
    elseif trim(Sensor2(1)) = "Short" then
      wscript.echo "Sensor 2: Short"
    end if
    if trim(Sensor3(1)) <> "Open" and trim(Sensor3(1)) <> "Short" then
      wscript.echo "Sensor: " & Field(0)
      wscript.echo "s/n: " & left(trim(Sensor3(1)), 16)
      wscript.echo "Current Temp: " & trim(Field(1))
      wscript.echo "Hi Temp: " & trim(Field(4))
      wscript.echo "Lo Temp: " & trim(Field(7))
      wscript.echo "Units: " & trim(Sensor3(3))
      if mid(Sensor3(1), 18, 19) <> "00-OK-00 CRC Errors" then
        wscript.echo "Data Integrity: OK"
      else
        wscript.echo "Data Integrity: Bad"
      end if
    elseif trim(Sensor3(1)) = "Short" then
      wscript.echo "Sensor 3: Short"
    end if
    if trim(Sensor4(1)) <> "Open" and trim(Sensor4(1)) <> "Short" then
      wscript.echo "Sensor: " & Field(0)
      wscript.echo "s/n: " & left(trim(Sensor4(1)), 16)
      wscript.echo "Current Temp: " & trim(Field(1))
      wscript.echo "Hi Temp: " & trim(Field(4))
      wscript.echo "Lo Temp: " & trim(Field(7))
      wscript.echo "Units: " & trim(Sensor4(3))
      if mid(Sensor4(1), 18, 19) <> "00-OK-00 CRC Errors" then
        wscript.echo "Data Integrity: OK"
      else
        wscript.echo "Data Integrity: Bad"
      end if
    elseif trim(Sensor4(1)) = "Short" then
      wscript.echo "Sensor 4: Short"
    end if

  end if
else
  wScript.echo "IsOpened = " & objComport.IsOpened
  wscript.echo "Error = " & objComport.LastError
  wScript.echo "Error description: " & objComport.GetErrorDescription( objComport.LastError )
end If

objComport.Close ()
