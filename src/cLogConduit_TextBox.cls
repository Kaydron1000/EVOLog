VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_TextBox"
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
Dim pTextBox As MSForms.TextBox

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
    pParentLoggers.Add ParentLogger
    pConduitName = ConduitName
    Me.ConduitLoggingLevel = ConduitLoggingLevel
End Sub
Sub InitTextBox(TextBox As MSForms.TextBox)
    Set pTextBox = TextBox
End Sub

Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
    If LogMessageLevel <> LoggingLevels.NoAssignment Then
        If LogMessageLevel >= Me.ConduitLoggingLevel Then
            pTextBox.Text = pTextBox.Text & FormatMessage(pMessageTemplate, LogMessagelLevel, Message) & vbNewLine
        End If
    End If
End Sub
Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
    Dim cnt As Long
    Dim sendStrg As String
    Dim sb As cStringBuilder
    
    Set sb = New cStringBuilder
    If UBound(LogMessageLevels) + 1 < LBound(LogMessageLevels) + EntryCount And UBound(Messages) + 1 < LBound(Messages) + EntryCount Then Err.Raise 440
    
    sb.Append pTextBox.Text
    For cnt = LBound(LogMessageLevels) To LBound(LogMessageLevels) + (EntryCount - 1)
        If LogMessageLevels(cnt) <> LoggingLevels.NoAssignment Then
            If LogMessageLevels(cnt) >= Me.ConduitLoggingLevel Then
                sb.Append FormatMessage(pMessageTemplate, LogMessageLevels(cnt), Messages(cnt)) & vbNewLine
            End If
        End If
    Next
    
    pTextBox.Text = sb.toString
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


