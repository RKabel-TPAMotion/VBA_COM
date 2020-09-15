Option Explicit
    Dim CLR As New CLRS232
    Dim status As Integer
        
    Public Function WriteReadCOM() As String
        Dim startString, writeString As String
        Dim dataArray
        
        CLR.FlushComms
        
        'Write
        
        startString = "@" & Sheets("Sheet1").Range("J13").value & ":"
        writeString = Sheets("Sheet1").Range("K13").value
        
        CLR.WriteComm (startString & writeString & vbCrLf)
        If CLR.status = 12 Then
            Debug.Print ("Write: " & startString & writeString & vbCrLf)
            Sheets("Sheet1").Range("L7").value = "Ready to Read"
        Else
            Debug.Print ("Nothing Written")
            Debug.Print ("status: " & CLR.status)
            Debug.Print ("errorMsg: " & CLR.ErrorMsg)
            Debug.Print (CLR.data)
            Sheets("Sheet1").Range("L7").value = "Failed to Write"
        End If

        'Read Response
        CLR.ReadComm
        
        If CLR.data = vbNullString Then
            Debug.Print ("Contains a null string: ")
            Debug.Print (CLR.data)
        Else
            Debug.Print ("no null string: ")
            Debug.Print (CLR.data)
        End If
        
        If Not CLR.data = vbNullString Then
            'temp close here due to issue of not closing right after this.
            dataArray = Split(CLR.data, vbCrLf)
            dataArray = Split(dataArray(0), ":")
            
            If dataArray(1) = "COMERR2" Then
                Sheets("Sheet1").Range("L7").value = "Invalid Input"
                Exit Function
            End If
            
            If dataArray(1) = "COMERR3" Then
                Sheets("Sheet1").Range("L7").value = "Invalid Input"
                Exit Function
            End If
            
            'looks like I need to do a check to see if "," is in there as well, and if an error does occur, I know
            'that the "," character is not actually in there, then continues on through the rest of the code here.
            
            dataArray = Split(dataArray(1), "=")

            Sheets("Sheet1").Range("K17") = CLR.data
            Sheets("Sheet1").Range("L7").value = "Ready to Write"
            'K20 is the item of the read data
            Sheets("Sheet1").Range("K20").value = dataArray(0)
            'L20 is the value of the read data
            Sheets("Sheet1").Range("L20").value = dataArray(1)
            'Debug.Print("Parsed Data: " & )
        End If
        
        
        
    End Function
            
    Sub test1()
        Dim data As String
        Dim dataArray
        
        data = "#01:VX=0" & vbCrLf
        
        dataArray = Split(data, vbCrLf)
        dataArray = Split(dataArray(0), ":")
        dataArray = Split(dataArray(1), "=")
                
        Debug.Print ("item: " & dataArray(0))
        Debug.Print ("value: " & dataArray(1))
        
    End Sub

    
    Sub WriteComBtn()
        Dim returnData As String
        returnData = WriteReadCOM()
        'Debug.Print ("ReturnData: " & returnData)
        'this returnData doesn't do anything, kind of no point in making a function of it, other than to be able to call it elsewhere
        'w/o having to press the button to call it.
     
    End Sub
   
    Sub ConnectToSerialPortBtn()

        Dim lngComPort, lngBaudRate, lngDataBits, lngStopBits, lngCol As Long
        Dim strParity As String
        
        With Sheets("Sheet1")
            lngComPort = .Range("P2").value
            lngBaudRate = .Range("P3").value
            strParity = .Range("P4").value
            lngDataBits = .Range("P5").value
            lngStopBits = .Range("P6").value
        End With
        
        'Debug.Print ("status: " & CLR.status)
        'Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        'Debug.Print ("LineDTR: " & CLR.LineDTR)
        'Debug.Print ("LineRTS: " & CLR.LineRTS)
        'Debug.Print ("data: " & CLR.Data)
        
        With CLR
            .COMport = lngComPort
            .BaudRate = lngBaudRate
            .Parity = strParity
            .Databits = lngDataBits
            .StopBits = lngStopBits
            .PostCommDelay = 0.1
            .OpenComms
        End With
        
        'The status is 5 and errorMsg is 5 since port is already open! So that is good to know it works
        If CLR.status <> 5 Then
            Sheets("Sheet1").Range("L7").value = "Open"
        Else
            MsgBox ("Port Already Open!")
            CLR.FlushComms
            CLR.CloseComms
            CLR.SerialConnectRetry
            CLR.OpenComms
        End If
        
    End Sub
    
    Sub DisconnectFromSerialPortBtn()
    'having an issue closing the comms for some reason
    'looks like I have to flush the Comms first in order to close the comms correctly!
    
        CLR.FlushComms
        CLR.CloseComms
        DoEvents
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("data: " & CLR.data)
        
        Sheets("Sheet1").Range("L7").value = "Closed"
        
    End Sub
