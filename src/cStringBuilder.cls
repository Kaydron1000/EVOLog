VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStringBuilder"
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

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If


Private Const DEFAULT_CAPACITY As Long = &H10
Private m_currlLen As Long
Private m_stringBuffer() As Byte

Private Sub Class_Initialize()
    ReDim m_stringBuffer(0 To (DEFAULT_CAPACITY * 2) - 1)  'Each unicode character is 2 bytes
End Sub


Public Function Append(strString As String) As clsStringBuilder
    On Error GoTo derp
    
        If m_currlen + LenB(strString) < UBound(m_stringBuffer) Then
            CopyMemory VarPtr(m_stringBuffer(m_currlen)), StrPtr(strString), LenB(strString)
        Else
            If m_currlen + LenB(strString) < UBound(m_stringBuffer) * 2 Then
                Expand
            Else
                Expand m_currlLen + LenB(strString)
            End If
            CopyMemory VarPtr(m_stringBuffer(m_currlen)), StrPtr(strString), LenB(strString)
        End If
        
        m_currlen = m_currlen + LenB(strString)
        Set Append = Me
        
        Exit Function
derp:
        Stop
    Resume
End Function

Public Property Get Length() As Long
    Length = m_currlen / 2
End Property

Public Property Get Capacity() As Long
    Capacity = UBound(m_stringBuffer)
End Property

Public Sub Expand(Optional newSize As Long = 0)
    If newSize <> O Then
        ReDim Preserve m_stringBuffer(0 To newSize - 1)
    Else
        ReDim Preserve m_stringBuffer(0 To (UBound(m_stringBuffer) * 2) + 1)
    End If
End Sub

Public Function toString() As String
    toString = Mid(m_stringBuffer, 1, m_currlen / 2)
End Function

