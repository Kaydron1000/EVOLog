VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_EvoLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ILogConduit
Dim pEvoLogger As cEvoLogger

Dim pParentLoggers As Collection
Dim pLoggingLevelNames As Collection
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Public Property Get ParentLoggers() As cEvoLoggers
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
    pMessageTemplate = Template
End Property

Property Get MessageTemplate() As String
    MessageTemplate = pMessageTemplate
End Property

Sub InitConduit(ParentLogger As cEvoLogger, ConduitName As String, ConduitLoggingLevel As LoggingLevels)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
    pConduitName = ConduitName
    Me.ConduitLoggingLevel = ConduitLoggingLevel
End Sub
Sub InitEvoLogger(EvoLogger As cEvoLogger)
    Set pEvoLogger = EvoLogger
End Sub
Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    pEvoLogger.LogEntry LogMessageLevel, Message
End Sub
Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    
End Sub
Private Function FormatMessage(MessageTemplate As String, LoggingLevel As LoggingLevels, Message As String)
    FormatMessage = UCase(MessageTemplate)
    FormatMessage = Replace(FormatMessage, "#LOGGINGLEVEL#", pLoggingLevelNames(LoggingLevel))
    FormatMessage = Replace(FormatMessage, "#DATE#", DateTime.Date)
    FormatMessage = Replace(FormatMessage, "#TIME#", DateTime.Time)
    FormatMessage = Replace(FormatMessage, "#TIMESTAMP#", DateTime.Now)
    FormatMessage = Replace(FormatMessage, "#MESSAGE#", Message)
End Function
Private Sub Class_Initialize()
    pMessageTemplate = "[$#LOGGINGLEVEL#] [#DATE# $TIME#]: $MESSAGE#"
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



