VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_MemoryLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
''This log conduit channels LogArtifact to an array in memory.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ILogConduit

Dim pParentLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pLogArtifactsArray() As cEvoArtifact
Dim pMemorySize As Integer
Dim pNextMemoryIndx As Integer
Dim pLoggingLevelNames As Collection

'*******************************************************************************
' Parent Loggers that log messages are channeled from.
'*******************************************************************************
Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLogger = pParentLogger
End Property

'*******************************************************************************
' ILogConduit Name
'*******************************************************************************
Property Get ConduitName() As String
    ConduitName = pConduitName
End Property

'*******************************************************************************
' Set the logging level that is channeled from this conduit.
'*******************************************************************************
Property Let ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = Level
End Property

'*******************************************************************************
' Get the logging level that is channeled from this conduit.
'*******************************************************************************
Property Get ConduitLoggingLevel() As LoggingLevels
    ConduitLoggingLevel = pConduitLoggingLevel
End Property

'*******************************************************************************
' Set the message template that describes the format of the log message.
'*******************************************************************************
Property Let MessageTemplate(Template As String)
    pMessageTemplate = Template
End Property

'*******************************************************************************
' Get the message template that describes the format of the log message.
'*******************************************************************************
Property Get MessageTemplate() As String
    MessageTemplate = pMessageTemplate
End Property

'*******************************************************************************
'
'*******************************************************************************
Property Let MemorySize(Size As Integer)
    Dim prevMemSize As Integer
    Dim indx As Integer
    
    If pNextMemoryIndx > Size Then Err.Raise 447
    
    prevMemSize = pMemorySize
    pMemorySize = Size
    ReDim Preserve pLogArtifactsArray(0 To MemorySize - 1)
    For indx = prevMemSize To pMemorySize - 1
        Set pLogArtifactsArray(indx) = New cEvoArtifact
    Next
End Property

'*******************************************************************************
'
'*******************************************************************************
Property Get MemorySize() As Integer
    MemorySize = pMemorySize
End Property




'*******************************************************************************
' Initializes properties of the conduit. Meant for when adding to a cEvoLogger.
'*******************************************************************************
Private Sub InitConduit(ParentLogger As cEvoLogger)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
End Sub

'*******************************************************************************
' Initializes properties of the conduit.
'*******************************************************************************
Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Verbose, Optional InitialMemorySize As Integer = 20)
    pConduitName = ConduitName
    Me.ConduitLoggingLevel = ConduitLoggingLevel
    MemorySize = InitialMemorySize
    pNextMemoryIndx = 0
End Sub

'*******************************************************************************
' Channels a single LogArtifact to its destination.
'*******************************************************************************
Private Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
    If LogArtifact.LoggingLevel >= Me.ConduitLoggingLevel Then
        If pNextMemoryIndx = pMemorySize Then MemorySize = MemorySize * 2
          
        pLogArtifactsArray(pNextMemoryIndx).Charged = True
        pLogArtifactsArray(pNextMemoryIndx).LoggingLevel = LogArtifact.LoggingLevel
        pLogArtifactsArray(pNextMemoryIndx).Message = LogArtifact.Message
        pLogArtifactsArray(pNextMemoryIndx).TimeStamp = LogArtifact.TimeStamp
        pNextMemoryIndx = pNextMemoryIndx + 1
    End If
End Sub

'*******************************************************************************
' Channels an array of LogArtifacts to their destinations.
'*******************************************************************************
Private Sub ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    Dim cnt As Long
    
    For cnt = LBound(LogArtifacts) To UBound(LogArtifacts)
        ChannelLogArtifact LogArtifacts(cnt)
    Next
End Sub

Public Sub ReChannelLogArtifacts(EvoLogger As cEvoLogger)
    Dim indx As Integer
    
    For indx = LBound(pLogArtifactsArray) To pNextMemoryIndx - 1
        EvoLogger.LogArtifactObject pLogArtifactsArray(indx)
    Next
End Sub

Private Sub Class_Initialize()
    Set pParentLoggers = New cEvoLoggers
End Sub


'*******************************************************************************
' Implementing ILogConduit Interface Methods and Properties
'*******************************************************************************


Private Property Get ILogConduit_ParentEvoLoggers() As cEvoLoggers
    Set ILogConduit_ParentEvoLoggers = Me.ParentEvoLoggers
End Property

Private Property Get ILogConduit_ConduitName() As String
    ILogConduit_ConduitName = Me.ConduitName
End Property

Private Property Let ILogConduit_ConduitLoggingLevel(Level As LoggingLevels)
    pConduitLoggingLevel = ILogConduit_ConduitLoggingLevel
End Property

Private Property Get ILogConduit_ConduitLoggingLevel() As LoggingLevels
    ILogConduit_ConduitLoggingLevel = pConduitLoggingLevel
End Property

Private Property Let ILogConduit_MessageTemplate(Template As String)
    pMessageTemplate = ILogConduit_MessageTemplate
End Property

Private Property Get ILogConduit_MessageTemplate() As String
    ILogConduit_MessageTemplate = pMessageTemplate
End Property

Private Sub ILogConduit_InitConduit(ParentLogger As cEvoLogger)
    InitConduit ParentLogger
End Sub

Private Sub ILogConduit_ChannelLogArtifact(LogArtifact As cEvoArtifact)
    ChannelLogArtifact LogArtifact
End Sub

Private Sub ILogConduit_ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    ChannelLogArtifacts LogArtifacts
End Sub



