Attribute VB_Name = "EvoLog"
Option Explicit
Public Enum EvoObjects
    EvoLogger
    LogConduit_Counter
    LogConduit_EvoLogger
    LogConduit_ExcelWorksheet
    LogConduit_File
    LogConduit_Immediate
    LogConduit_MemoryLogger
    LogConduit_TextBox
End Enum
Function CreateEvo(EvoObject As EvoObjects) As Variant
    Select Case EvoObject
        Case EvoLogger
            Set CreateEvo = New cEvoLogger
        Case LogConduit_Counter
            Set CreateEvo = New cLogConduit_Counter
        Case LogConduit_EvoLogger
            Set CreateEvo = New cLogConduit_EvoLogger
        Case LogConduit_ExcelWorksheet
            Set CreateEvo = New cLogConduit_ExcelWorksheet
        Case LogConduit_File
            Set CreateEvo = New cLogConduit_File
        Case LogConduit_Immediate
            Set CreateEvo = New cLogConduit_Immediate
        Case LogConduit_MemoryLogger
            Set CreateEvo = New cLogConduit_MemoryLogger
        Case LogConduit_TextBox
            Set CreateEvo = New cLogConduit_TextBox
    End Select
End Function
