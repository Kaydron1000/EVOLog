VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogConduit"
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
''EVOLogger has the capablity to interface with any ILogConduit
''created. This conduit definition provides the neccessary
''properties and subroutines for cEVOLogger class to utilze a
''conduit. Conduits provide a way to channel a log message to
''its appropriate destination (i.e. Text File,Text Box).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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

Sub InitConduit(ParentLogger As cEvoLogger)
End Sub

Sub ChannelLogArtifact(LogArtifact As cEvoArtifact)
End Sub

Sub ChannelLogArtifacts(LogArtifact() As cEvoArtifact)
End Sub


