Attribute VB_Name = "TestSamples"
Sub tsts()
    Dim EvoLogger As cEvoLogger
    Dim logCounter As cLogConduit_Counter
    Dim logtb  As cLogConduit_TextBox
    Dim b As String
    Dim coll As Collection
    
    Set coll = New Collection
    coll.Add "A"
    coll.Add "B"
    coll.Add "C"
    coll.Add "D"
    
    coll.Remove 1
    
    a = coll.Count
    
    coll.Remove 1
    coll.Remove 1
    coll.Remove 1
    
    a = coll.Count
    
    Set EvoLogger = New cEvoLogger
    Set logCounter = New cLogConduit_Counter
    Set logtb = New cLogConduit_TextBox
    b = "astrg"
    r = TypeName(logCounter)
    If TypeOf logCounter Is ILogConduit Then
        a = a
    End If
    If TypeOf logCounter Is cLogConduit_Counter Then
        a = a
    End If
    r = TypeName(b)
    logCounter.InitConduit EvoLogger, "Name", LoggingLevels.Error
   
    
            
    EvoLogger.AddConduit logCounter
    
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    EvoLogger.LogEntry Error, "new mess"
    
    EvoLogger.FlushBatchedLogEntries
    Debug.Print logCounter.LogCountsToString
End Sub
Sub UnitTest_cEVOLogger_()
    Dim PauseTime, Start, Finish, TotalTime
    Dim s As String
    Dim strg As String
    Dim itms() As Double
    ReDim itms(1 To 5)
'    s = Format(Now, "yyyy-mm-dd hh:mm:ss") & Right(Format(Timer, "0.000"), 4)
'    Debug.Print s
'    If (MsgBox("Press Yes to pause for 5 seconds", 4)) = vbYes Then
'        PauseTime = 5    ' Set duration.
'        Start = Timer    ' Set start time.
'        Do While Timer < Start + PauseTime
'            DoEvents    ' Yield to other processes.
'        Loop
'        Finish = Timer    ' Set end time.
'        TotalTime = Finish - Start    ' Calculate total time.
'        MsgBox "Paused for " & TotalTime & " seconds"
'    Else
'        End
'    End If
    strg = "My Really long string that i will take one char at a time to put into a gaint array. ... i need to put this in a var to get length."

    arr = Split(strg, "")
    For r = 1 To 5
        Start = Timer
        For n = 1 To 1000000
            For q = 1 To Len(strg)
                s = Mid(strg, q, 1)
            Next
        Next
        Finish = Timer
        itms(r) = Finish - Start
    Next
    For r = 1 To 5
        mysum = itms(r) + mysum
    Next
    Debug.Print "---- mid"
    Debug.Print mysum / 5
End Sub
