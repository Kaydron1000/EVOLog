VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogConduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get ParentEvoLoggers() As cEvoLoggers
End Property

Property Get ConduitName() As String
End Property

Property Let ConduitLoggingLevel(Level As LoggingLevels)
End Property

Property Get ConduitLoggingLevel() As LoggingLevels
End Property

Property Let MessageTemplate(Template As String)
End Property

Property Get MessageTemplate() As String
End Property

Sub InitConduit(ParentLogger As cEvoLogger, ConduitName As String, ConduitLoggingLevel As LoggingLevels)
End Sub

Sub ChannelLogEntry(LogMessageLevel As LoggingLevels, Message As String)
End Sub

Sub ChannelLogEntries(LogMessageLevels() As LoggingLevels, Messages() As String, EntryCount As Integer)
End Sub

