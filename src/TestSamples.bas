Attribute VB_Name = "TestSamples"
Sub tsts()
    Dim EvoLogger As cEvoLogger
    Dim logCounter As cLogConduit_Counter
    Dim logtb  As cLogConduit_TextBox
    
    Set EvoLogger = New cEvoLogger
    Set logCounter = New cLogConduit_Counter
    Set logtb = New cLogConduit_TextBox
    
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
