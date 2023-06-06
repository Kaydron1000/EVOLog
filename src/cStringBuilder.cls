VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

