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

End Sub
