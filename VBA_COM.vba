Option Explicit
    Dim btnWStart, btnWStop, btnWFlag As Boolean
    Dim btnRStart, btnRStop, btnRFlag As Boolean
    Dim btnStartCom, btnCloseCom As Boolean
    Dim CLR As New CLRS232
    Dim status As Integer
    
    
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
        CLR.WriteComm (startString & writeString & vbCrLf)
        Debug.Print (startString & writeString & vbCrLf)
        
        'thanks to this status and error msg below it is showing that nothing was every sent!
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
    End Sub

    Sub ReadComBtn()
        Dim ReadString As String
        Dim byte1 As Byte, chars As String
        
        chars = ""
        
        Debug.Print ("Read: ")
        CLR.ReadComm
        'DoEvents
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
        Debug.Print ("byte1: " & byte1)
        
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
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
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
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
    End Sub

    Sub ReConnectToSerialPort()
        Dim data, data2 As String
        Dim dataArray, dataArray2
        'CLR.SerialConnectRetry
        data = "#01:VX=0" & vbCrLf
        Debug.Print ("data: " & data)
        dataArray = Split(data, vbCrLf)
        Debug.Print ("dataArray: " & dataArray(0))
        data2 = dataArray(0)
        'removes the vbCrLf to use it, now just to split off at the ':' portion
        dataArray2 = Split(dataArray(0), ":")
        Debug.Print ("data2: " & dataArray2(1))
    'works, now just have to remove some variables to where I would like it instead now
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
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
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
            Sheets("Sheet1").Range("L7").Value = "Open"
        Else
            MsgBox ("Port Already Open, can't Open right now!")
        End If
        
    End Sub
    
    Sub DisconnectFromSerialPortBtn()
        CLR.CloseComms
        
        Debug.Print ("status: " & CLR.status)
        Debug.Print ("errorMsg: " & CLR.ErrorMsg)
        Debug.Print ("LineDTR: " & CLR.LineDTR)
        Debug.Print ("LineRTS: " & CLR.LineRTS)
        Debug.Print ("data: " & CLR.Data)
        
        btnCloseCom = True
        btnStartCom = False
        
        Sheets("Sheet1").Range("L7").Value = "Closed"
        
    End Sub
