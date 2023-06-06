Attribute VB_Name = "TestSamples"
Sub tsts()
    Dim evoLogger As cEvoLogger
    Dim logCounter As cLogConduit_Counter
    Dim logtb  As cLogConduit_TextBox
    
    Set evoLogger = New cEvoLogger
    Set logCounter = New cLogConduit_Counter
    Set logtb = New cLogConduit_TextBox
    
    logCounter.InitConduit evoLogger, "Name", LoggingLevels.Error
    
    evoLogger.AddConduit logCounter
    
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    evoLogger.LogEntry Error, "new mess"
    
    evoLogger.FlushBatchedLogEntries
    Debug.Print logCounter.LogCountsToString
End Sub
