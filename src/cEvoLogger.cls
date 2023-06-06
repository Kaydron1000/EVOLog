VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvoLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''
'''' Logger class was inspired by Serilog library for .NET.
'''' Provides ability for multiple conduits
''''''''''''''''''''''''''
Option Explicit
Public Enum LoggingLevels
    NoAssignment = 0
    Verbose = 1
    Debugg = 2
    Information = 3
    Warning = 4
    Error = 5
    Fatal = 6
End Enum
Enum FileAccesses
    Overwrite
    Append
    CreateNew
End Enum

Private pLoggerName As String
Private pActiveConduits As cLogConduits
Private pLoggingLevelNames As Collection
Private pBatchedLoggingLevels() As LoggingLevels
Private pBatchedLogMessages() As String
Private pBatchSetCount As Integer
Private pBatchOutput As Boolean
Private pNextBatchCnt As Integer

Public Property Get LoggerName() As String
    LoggerName = pLoggerName
End Property
Public Property Get BatchOutput() As Boolean
    BatchOutput = pBatchOutput
End Property
Public Property Let BatchOutput(batchIt As Boolean)
    If batchIt Then
        ReDim pBatchedLoggingLevels(0 To pBatchSetCount - 1)
        ReDim pBatchedLogMessages(0 To pBatchSetCount - 1)
    Else
        FlushBatchedLogEntries
        Erase pBatchedLoggingLevels
        Erase pBatchedLogMessages
    End If
    
    pBatchOutput = batchIt
End Property
Public Property Get BatchSetCount() As Integer
    BatchSetCount = pBatchSetCount
End Property
Public Property Let BatchSetCount(BatchCount As Integer)
    pBatchSetCount = BatchCount
    ReDim Preserve pBatchedLoggingLevels(0 To pBatchSetCount - 1)
    ReDim Preserve pBatchedLogMessages(0 To pBatchSetCount - 1)
End Property
Public Property Get ConduitsCount() As Integer
    ConduitsCount = pActiveConduits.Count
End Property
Public Property Get LoggingLevelNames() As Collection
    Set LoggingLevelNames = pLoggingLevelNames
End Property
Sub AddConduit(Conduit As ILogConduit)
    pActiveConduits.Add Conduit, Conduit.ConduitName
End Sub
Sub RemoveConduit(LogConduitName As String)
    pActiveConduits.Remove LogConduitName
End Sub
Function GetConduit(Item As Variant) As ILogConduit
    Set GetConduit = pActiveConduits(Item)
End Function
Sub ClearConduits()
    Dim cnt As Integer
    
    For cnt = 1 To pActiveConduits.Count
        pActiveConduits.Remove cnt
    Next
End Sub
Sub LogEntry(LoggingLevel As LoggingLevels, LogMessage As String)
    Dim logConduit As ILogConduit
    
    If pBatchOutput Then
        pBatchedLoggingLevels(pNextBatchCnt) = LoggingLevel
        pBatchedLogMessages(pNextBatchCnt) = LogMessage
        
        pNextBatchCnt = pNextBatchCnt + 1
        
        If pNextBatchCnt >= pBatchSetCount Then
            FlushBatchedLogEntries
        End If
    Else
        For Each logConduit In pActiveConduits
            logConduit.ChannelLogEntry LoggingLevel, LogMessage
        Next
    End If
End Sub

Sub FlushBatchedLogEntries()
    Dim logConduit As ILogConduit
    Dim cnt As Integer
    
    If BatchOutput Then
        For Each logConduit In pActiveConduits
            logConduit.ChannelLogEntries pBatchedLoggingLevels, pBatchedLogMessages, pNextBatchCnt
        Next
        
        For cnt = 0 To pBatchSetCount - 1
            pBatchedLoggingLevels(cnt) = LoggingLevels.NoAssignment
            pBatchedLogMessages(cnt) = ""
        Next
        
        pNextBatchCnt = 0
    End If
End Sub
Private Function GUID(Optional lowercase As Boolean, Optional parens As Boolean) As String
    Dim k&, h$
    GUID = Space(36)
    For k = 1 To Len(GUID)
        Randomize
        Select Case k
            Case 9, 14, 19, 24: h = "-"
            Case 15:            h = "4"
            Case 20:            h = Hex(Rnd * 3 + 8)
            Case Else:          h = Hex(Rnd * 15)
        End Select
        Mid$(GUID, k, 1) = h
    Next
    If lowercase Then GUID = LCase$(GUID)
    If parens Then GUID = "{" & GUID & "}"
End Function
Private Sub Class_Initialize()
    Set pActiveConduits = New cLogConduits
    Set pLoggingLevelNames = New Collection
    pLoggingLevelNames.Add "Verbose", "Verbose"
    pLoggingLevelNames.Add "Debugg", "Debugg"
    pLoggingLevelNames.Add "Information", "Information"
    pLoggingLevelNames.Add "Warning", "Warning"
    pLoggingLevelNames.Add "Error", "Error"
    pLoggingLevelNames.Add "Fatal", "Fatal"
    pBatchSetCount = 20
    ReDim pBatchedLoggingLevels(0 To pBatchSetCount - 1)
    ReDim pBatchedLogMessages(0 To pBatchSetCount - 1)
    pBatchOutput = True
    pLoggerName = GUID()
End Sub


