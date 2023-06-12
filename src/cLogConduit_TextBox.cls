VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLogConduit_TextBox"
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
''This log conduit channels log messages to a MSForms.Textbox.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Implements ILogConduit

Dim pParentLoggers As cEvoLoggers
Dim pConduitName As String
Dim pConduitLoggingLevel As LoggingLevels
Dim pMessageTemplate As String

Dim pTextBox As MSForms.TextBox

'*******************************************************************************
' Parent Loggers that log messages are channeled from.
'*******************************************************************************
Public Property Get ParentLoggers() As cEvoLoggers
    Set ParentLoggers = pParentLoggers
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
' Initializes properties of the conduit. Meant for when adding to a cEvoLogger.
'*******************************************************************************
Private Sub InitConduit(ParentLogger As cEvoLogger)
    pParentLoggers.Add ParentLogger, ParentLogger.LoggerName
End Sub

'*******************************************************************************
' Initializes properties of the conduit. Initializes everything for the
' channeling to a Textbox except for the parent cEvoLogger object.
'*******************************************************************************
Sub Init(TextBox As MSForms.TextBox, ConduitName As String, Optional ConduitLoggingLevel As LoggingLevels = Information)
    Set pTextBox = TextBox
    Me.ConduitName = ConduitName
    
    Me.ConduitLoggingLevel = ConduitLoggingLevel
End Sub

'*******************************************************************************
' Channels a single LogArtifact to its destination.
'*******************************************************************************
Private Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
    If LogArtifact.LoggingLevel >= Me.ConduitLoggingLevel Then
        pTextBox.Text = pTextBox.Text & LogArtifact.FormatArtifact(pMessageTemplate) & vbNewLine
    End If
End Sub

'*******************************************************************************
' Channels an array of LogArtifacts to their destinations.
'*******************************************************************************
Private Sub ChannelLogArtifacts(LogArtifacts() As cEvoArtifact)
    Dim cnt As Long
    Dim sendStrg As String
    Dim sb As cStringBuilder
    
    Set sb = New cStringBuilder
    
    sb.Append pTextBox.Text
    For cnt = LBound(LogArtifacts) To UBound(LogArtifacts)
        If LogArtifacts(cnt).LoggingLevel >= Me.ConduitLoggingLevel Then
            sb.Append LogArtifacts(cnt).FormatArtifact(pMessageTemplate) & vbNewLine
        End If
    Next
    
    pTextBox.Text = sb.ToString
End Sub

'*******************************************************************************
' Class initialize method
'*******************************************************************************
Private Sub Class_Initialize()
    pMessageTemplate = "[$#LOGGINGLEVEL#] [#DATE# $TIME#]: $MESSAGE#"
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
