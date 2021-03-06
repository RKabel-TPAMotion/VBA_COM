Option Explicit
    Dim btnWStart, btnWStop, btnWFlag As Boolean
    Dim btnRStart, btnRStop, btnRFlag As Boolean
    Dim btnStartCom, btnCloseCom As Boolean
    Dim CLR As New CLRS232
    
    
    Public Function WriteCom() As String
        
    End Function
    
    Sub WriteComBtn()
        'ex:
        '@01:ex=200 or @01:EX=200 or @01:EX;VX or @01:EX;VX=200
        'want to take the value in cell L13 and then add the vbCrLf to the end before submitting to the writecomms function
        
        Dim startString As String
        Dim writeString As String
        
        startString = "@01:"
        writeString = Sheets("Sheet1").Range("K13").Value
        
        'CLR.FlushComms
        
        Debug.Print ("Write: ")
        CLR.WriteComm (startString & writeString)
        Debug.Print (startString & writeString)
        
    End Sub

    Sub ReadComBtn()
        Dim ReadString As String
        Dim byte1 As Byte, chars As String
        
        chars = ""
        
        Debug.Print ("Read: ")
        CLR.ReadComm
        'DoEvents
        
        If byte1 = Chr(13) Then
            Sheets("Sheet1").Range("K17").Value = chars
            Debug.Print ("chars: ")
            Debug.Print (chars)
            chars = ""
        Else
            chars = chars & Chr(byte1)
            Debug.Print ("else - chars: ")
            Debug.Print (chars)
        End If
        
        If Not CLR.Data = vbNullString Then
            Sheets("Sheet1").Range("K17") = CLR.Data
            'don't know if above or below one works yet, will need to do further testing
            'Sheets("Sheet1").Range("K17").Value = CLR.Data
            Sheets("Sheet1").Range("L7").Value = "Transfer"
        End If
        
        If CLR.Data = vbNullString Then
            Debug.Print ("Contains a null string: ")
            Debug.Print (CLR.Data)
        Else
            Debug.Print ("no null string: ")
            Debug.Print (CLR.Data)
        End If
        
    End Sub
    
    Sub ConnectToSerialPortBtn()

        Dim lngComPort, lngBaudRate, lngDataBits, lngStopBits, lngCol As Long
        Dim strParity As String
        
        btnStartCom = True
        btnCloseCom = False
        
        With Sheets("Sheet1")
            lngComPort = .Range("P2").Value
            lngBaudRate = .Range("P3").Value
            strParity = .Range("P4").Value
            lngDataBits = .Range("P5").Value
            lngStopBits = .Range("P6").Value
        End With
        
        With CLR
            .COMport = lngComPort
            .BaudRate = lngBaudRate
            .Parity = strParity
            .Databits = lngDataBits
            .StopBits = lngStopBits
            .PostCommDelay = 0.1
            .OpenComms
        End With
        
        Sheets("Sheet1").Range("L7").Value = "Open"
        
    End Sub
    
    Sub DisconnectFromSerialPortBtn()
        CLR.CloseComms
        
        btnCloseCom = True
        btnStartCom = False
        
        Sheets("Sheet1").Range("L7").Value = "Closed"
        
    End Sub
