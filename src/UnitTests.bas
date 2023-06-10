Attribute VB_Name = "UnitTests"
Option Explicit
Sub UnitTest_cEvoLogger_Init()
    Dim logger As cEvoLogger
    Dim cnt As Integer
    Dim strg As String
    Dim errNum As Integer
    Dim coll As Collection
    Dim ConduitObj As ILogConduit
    
    Dim logCounter As cLogConduit_Counter
    
    Set logger = New cEvoLogger
    Set logCounter = New cLogConduit_Counter
    
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

