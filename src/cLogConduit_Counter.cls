VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ILogConduit

Dim pParentLoggers As cEvoLoggers
Dim pLoggingLevelNames As Collection
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String
Dim pLogCounts() As Long


Property Get ParentEvoLoggers() As cEvoLoggers
    Set ParentLoggers = pParentLoggers
End Property

Property Get ConduitName() As String
    ConduitName = pConduitName
End Property

Property Let ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = Level
End Property

Property Get ConduitLoggingLevel() As LoggingLevels
    ConduitLoggingLevel = pConduitLoggingLevel
End Property

Property Let MessageTemplate(Template As String)
End Property

Property Get MessageTemplate() As String
End Property

Property Get VerboseCount() As Long
    VerboseCount = pLogCounts(1)
End Property

Property Get DebuggCount() As Long
    DebuggCount = pLogCounts(2)
End Property

Property Get InformationCount() As Long
    InformationCount = pLogCounts(3)
End Property

Property Get WarningCount() As Long
    WarningCount = pLogCounts(4)
End Property
Property Get ErrorCount() As Long
    ErrorCount = pLogCounts(5)
End Property
Property Get FatalCount() As Long
    FatalCount = pLogCounts(6)
End Property
Sub InitConduit(ParentLogger As cEvoLogger, ConduitName As String, ConduitLoggingLevel As LoggingLevels)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
    pConduitName = ConduitName
    Me.ConduitLoggingLevel = ConduitLoggingLevel
    
    ReDim pLogCounts(1 To ParentLogger.LoggingLevelNames.Count)
End Sub
Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    If LogMessageLevel >= Me.ConduitLoggingLevel Then
        pLogCounts(LogMessageLevel) = pLogCounts(LogMessageLevel) + 1
    End If
End Sub

Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Dim cnt As Long
    
    If UBound(LogMessageLevels) + 1 < (LBound(LogMessageLevels) + EntryCount) And _
    UBound(Messages) + 1 < (LBound(Messages) + EntryCount) Then Err.Raise 440
    
    For cnt = LBound(LogMessageLevels) To LBound(LogMessageLevels) + (EntryCount - 1)
        If LogMessageLevels(cnt) <> LoggingLevels.NoAssignment Then
            If LogMessageLevels(cnt) >= Me.ConduitLoggingLevel Then
                pLogCounts(LogMessageLevels(cnt)) = pLogCounts(LogMessageLevels(cnt)) + 1
            End If
        End If
    Next
End Sub

Function LogCountsToString() As String
    Dim entries() As String
    Dim strg As String
    Dim n As Integer
    
    strg = ""
    
    ReDim entries(1 To pLoggingLevelNames.Count)
    
    For n = 1 To pLoggingLevelNames.Count
        If n >= Me.ConduitLoggingLevel Then
            entries(n) = pLoggingLevelNames(n) & " Count = " & pLogCounts(n)
        Else
            entries(n) = pLoggingLevelNames(n) & " Count = N/A"
        End If
    Next
    
    strg = Join(entries, ", ")

    
    LogCountsToString = strg
End Function

Sub ClearCounts()
    ReDim pLogCounts(1 To ParentLogger.LoggingLevelNames.Count)
End Sub

Private Sub Class_Initialize()
    Set pParentLoggers = New cEvoLoggers
    
    Set pLoggingLevelNames = New Collection
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
End Sub

Property Get ILogConduit_ParentEvoLoggers() As cEvoLoggers
    Set ILogConduit_ParentEvoLoggers = Me.ParentEvoLoggers
End Property

Property Get ILogConduit_ConduitName() As String
    ILogConduit_ConduitName = Me.ConduitName
End Property

Property Let ILogConduit_ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = ILogConduit_ConduitLoggingLevel
End Property

Property Get ILogConduit_ConduitLoggingLevel() As LoggingLevels
    ILogConduit_ConduitLoggingLevel = pConduitLoggingLevel
End Property

Property Let ILogConduit_MessageTemplate(Template As String)
    pMessageTemplate = ILogConduit_MessageTemplate
End Property

Property Get ILogConduit_MessageTemplate() As String
    ILogConduit_MessageTemplate = pMessageTemplate

End Property

Sub ILogConduit_InitConduit(ParentLogger As cEvoLogger, ConduitName As String, ConduitLoggingLevel As LoggingLevels)
    Me.InitConduit ParentLogger, ConduitName, ConduitLoggingLevel
End Sub

Sub ILogConduit_ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    Me.ChannelLogEntry LogMessageLevel, Message
End Sub

Sub ILogConduit_ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Me.ChannelLogEntries LogMessageLevels, Messages, EntryCount
End Sub


