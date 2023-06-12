Attribute VB_Name = "UnitTests"
Option Explicit
Sub UnitTest_cEvoLogger_Init()
    Dim logger As cEvoLogger
    Dim cnt As Integer
    Dim strg As String
    Dim errNum As Integer
    Dim coll As Collection
    Dim ConduitObj As ILogConduit
    
    Set logger = New cEvoLogger
    
    logger.Init "MyName"
    logger.BatchOutput = False
    logger.BatchOutput = True
    
    logger.BatchSetCount = 10
    logger.BatchSetCount = 20
    
    cnt = logger.ConduitsCount
    If cnt <> 0 Then Err.Raise -1
    
    strg = logger.LoggerName
    If strg <> "MyName" Then Err.Raise -1
    
    On Error Resume Next
    logger.FlushBatchedLogArtifacts
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise -1, "UnitTest_cEvoLogger_Init - logger.FlushBatchedLogArtifacts", "Error occured when it shouldn't"
    
    
    On Error Resume Next
    Set coll = logger.GetConduitNames
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise -1, "UnitTest_cEvoLogger_Init - logger.GetConduitNames", "Error occured when it shouldn't"
    
    
    On Error Resume Next
    logger.ClearConduits
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise -1, "UnitTest_cEvoLogger_Init - logger.ClearConduits", "Error occured when it shouldn't"
    
    
    On Error Resume Next
    Set ConduitObj = logger.GetConduit(1)
    errNum = Err.Number
    On Error GoTo 0
    If errNum = 0 Then Err.Raise -1, "UnitTest_cEvoLogger_Init - logger.GetConduit", "Error should have occcured, but it didn't"
    
    
    On Error Resume Next
    Set ConduitObj = logger.RemoveConduit(1)
    errNum = Err.Number
    On Error GoTo 0
    If errNum = 0 Then Err.Raise -1, "UnitTest_cEvoLogger_Init - logger.RemoveConduit", "Error should have occcured, but it didn't"
End Sub

Sub UnitTest_cLogConduit_Counter_Init()
    Dim logger As cEvoLogger
    Dim cnt As Variant
    Dim strg As String
    Dim errNum As Integer
    Dim coll As Collection
    Dim ConduitObj As ILogConduit
    Dim logCounter As cLogConduit_Counter
    
    Set logger = New cEvoLogger
    Set logCounter = New cLogConduit_Counter
    
    
    logger.Init "Logger"
    logCounter.Init "GBL_Counter"
    
    logger.AddConduit logCounter
    
    For cnt = 1 To logger.LoggingLevelNames.Count
        logger.LogArtifact CInt(cnt), logger.LoggingLevelNames(cnt) & " Message"
    Next
    logger.FlushBatchedLogArtifacts
    Debug.Print logCounter.LogCountsToString
End Sub
Sub UnitTest_cLogConduit_Immediate_Init()
    Dim logger As cEvoLogger
    Dim cnt As Variant
    Dim strg As String
    Dim errNum As Integer
    Dim coll As Collection
    Dim ConduitObj As ILogConduit
    Dim logCond As cLogConduit_Immediate
    
    Set logger = New cEvoLogger
    Set logCond = New cLogConduit_Immediate
    
    logCond.Init Verbose
    logger.Init "Logger"
    logger.BatchSetCount = 3
    
    logger.AddConduit logCond
    logger.LogArtifact Error, "This is an error"
    logger.LogArtifact Debugg, "This is debug"
    logger.LogArtifact Verbose, "This is Verbose"
    logger.LogArtifact Information, "This is info1"
    logger.LogArtifact Information, "This is info2"
    
    logger.BatchSetCount = 5
    
    logger.LogArtifact Verbose, "This is Verbose1"
    logger.LogArtifact Debugg, "This is debug1"
    logger.LogArtifact Error, "This is an error1"
    
    logger.LogArtifact Error, "This is an error t1"
    logger.LogArtifact Error, "This is an error t2"
    logger.LogArtifact Error, "This is an error t3"
    
    logger.BatchSetCount = 2
    
    logger.LogArtifact Error, "This is an INFO 1"
    logger.FlushBatchedLogArtifacts
End Sub
Sub UnitTest_cLogConduit_File_Init()
    Dim logger As cEvoLogger
    Dim cnt As Variant
    Dim strg As String
    Dim errNum As Integer
    Dim coll As Collection
    Dim ConduitObj As ILogConduit
    Dim logCond As cLogConduit_File
    
    Set logger = New cEvoLogger
    Set logCond = New cLogConduit_File
    
    
    logger.Init "EvoLogger"
    logger.BatchSetCount = 3
    
    logCond.Init "LogFile", Verbose
    logCond.OpenLogFile
    
    logger.AddConduit logCond
    logger.LogArtifact Error, "This is an error"
    logger.LogArtifact Debugg, "This is debug"
    logger.LogArtifact Verbose, "This is Verbose"
    logger.LogArtifact Information, "This is info1"
    logger.LogArtifact Information, "This is info2"
    Application.Wait Now + #12:00:01 AM#
    logger.BatchSetCount = 5
    
    logger.LogArtifact Verbose, "This is Verbose1"
    logger.LogArtifact Debugg, "This is debug1"
    logger.LogArtifact Error, "This is an error1"
    
    logger.LogArtifact Error, "This is an error t1"
    logger.LogArtifact Error, "This is an error t2"
    logger.LogArtifact Error, "This is an error t3"
    
    logger.BatchSetCount = 2
    
    logger.LogArtifact Error, "This is an INFO 1"
    logger.FlushBatchedLogArtifacts
    logCond.CloseLogFile
End Sub
