VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEvoLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''EVOLog - a logging platform for VBA inspired by Serilog Library for .NET.
''--------------------------------------
''https://github.com/Kaydron1000/EVOLog
''--------------------------------------
''
''Copyright (C) 2023  Brandon Neff
''
''SPDX-License-Identifier: GPL-2.0
''This program is free software; you can redistribute it and/or
''modify it under the terms of the GNU General Public License
''as published by the Free Software Foundation; either version 2
''of the License, or (at your option) any later version.
''
''This program is distributed in the hope that it will be useful,
''but WITHOUT ANY WARRANTY; without even the implied warranty of
''MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
''GNU General Public License for more details.
''
''You should have received a copy of the GNU General Public License
''along with this program; if not, write to the Free Software
''Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''EVOLogger's base class is the cEVOLogger. This class allows
''channeling log messages to multiple conduits. Once the class
''is instantiated conduits must be added to channel log
''messages out to the conduits destination. Conduits implement
''the ILogConduit interface to channel log messages to any
''destination (i.e. Text File,Text Box, another cEvoLogger).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Unique name of this logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LoggerName() As String
    LoggerName = pLoggerName
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Specifies if log messages are channeled in batches.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BatchOutput() As Boolean
    BatchOutput = pBatchOutput
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Sets log messages to be channeled individually or in batches.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Number of log messages stored before channeled in a batch.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get BatchSetCount() As Integer
    BatchSetCount = pBatchSetCount
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Set number of log messages to batch.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Let BatchSetCount(BatchCount As Integer)
    pBatchSetCount = BatchCount
    ReDim Preserve pBatchedLoggingLevels(0 To pBatchSetCount - 1)
    ReDim Preserve pBatchedLogMessages(0 To pBatchSetCount - 1)
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Number of attached conduits.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get ConduitsCount() As Integer
    ConduitsCount = pActiveConduits.Count
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Collection of the logging level names.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Property Get LoggingLevelNames() As Collection
    Set LoggingLevelNames = pLoggingLevelNames
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Adds a conduit to the logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub AddConduit(Conduit As ILogConduit)
    
    Me.FlushBatchedLogEntries
    Conduit.InitConduit Me
    pActiveConduits.Add Conduit, Conduit.ConduitName
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Removes a conduit from the logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub RemoveConduit(Conduit As Variant)
    Dim ConduitObj As ILogConduit
    
    If TypeName(Conduit) = "String" Or TypeName(Conduit) = "Integer" Then
        Set ConduitObj = pActiveConduits(Conduit)
    ElseIf TypeOf Conduit Is ILogConduit Then
        Set ConduitObj = Conduit
    Else
        Set ConduitObj = Nothing
    End If

    If Not ConduitObj Is Nothing Then
        Me.FlushBatchedLogEntries
        
        ConduitObj.ParentEvoLoggers.Remove Me.LoggerName
        pActiveConduits.Remove ConduitObj.ConduitName
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Gets the names of all conduits attached to this logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetConduitNames() As Collection
    Set GetConduitNames = pActiveConduits.Keys
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Gets the specified conduit.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetConduit(Item As Variant) As ILogConduit
    Set GetConduit = pActiveConduits(Item)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Clears all conduits for the logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ClearConduits()
    Dim cnt As Integer
    
    Do While pActiveConduits.Count > 0
        pActiveConduits.Remove cnt
    Loop
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Log a message to all attached conduits.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Flush any batched log messages to the conduits.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FlushBatchedLogEntries()
    Dim logConduit As ILogConduit
    Dim cnt As Integer
    
    If BatchOutput Then
        For Each logConduit In pActiveConduits
            logConduit.ChannelLogEntries pBatchedLoggingLevels, _
                                         pBatchedLogMessages, pNextBatchCnt
        Next
        
        For cnt = 0 To pBatchSetCount - 1
            pBatchedLoggingLevels(cnt) = LoggingLevels.NoAssignment
            pBatchedLogMessages(cnt) = ""
        Next
        
        pNextBatchCnt = 0
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Creates a moderately unique ID for logger.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GUID(Optional lowercase As Boolean, _
                      Optional parentheses As Boolean) As String
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
    If parentheses Then GUID = "{" & GUID & "}"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Initializes EVOLogger class.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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


