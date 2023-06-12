VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_Counter"
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
''This log conduit stores the count of log levels channeled
''through this conduit.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
Private Sub InitConduit(ParentLogger As cEvoLogger)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
    
    ReDim pLogCounts(1 To pLoggingLevelNames.Count)
End Sub

Sub Init(ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Verbose)
    pConduitName = ConduitName
    
    Me.ConduitLoggingLevel = ConduitLoggingLevel
End Sub
'*******************************************************************************
' Channels a single LogArtifact to its destination.
'*******************************************************************************
Private Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
    If LogArtifact.LoggingLevel >= Me.ConduitLoggingLevel Then
        pLogCounts(LogArtifact.LoggingLevel) = pLogCounts(LogArtifact.LoggingLevel) + 1
    End If
End Sub

'*******************************************************************************
' Channels an array of LogArtifacts to their destinations.
'*******************************************************************************
Private Sub ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    Dim cnt As Long
    
    For cnt = LBound(LogArtifacts) To UBound(LogArtifacts)
        If LogArtifacts(cnt).LoggingLevel >= Me.ConduitLoggingLevel Then
            pLogCounts(LogArtifacts(cnt).LoggingLevel) = pLogCounts(LogArtifacts(cnt).LoggingLevel) + 1
        End If
    Next
End Sub

Function LogCountsToString(Optional ClearCounts As Boolean = False) As String
    Dim LoggingLevelNames As Collection
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
    
    If ClearCounts Then ClearCounts
    
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

