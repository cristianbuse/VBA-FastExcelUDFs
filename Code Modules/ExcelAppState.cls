VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelAppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
'' VBA FastExcelUDFs
'''-------------------------------------------------
'' https://github.com/cristianbuse/VBA-FastExcelUDFs
'''-------------------------------------------------
'''
''' Copyright (c) 2019 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
'''=============================================================================
''
''==============================================================================
'' Description:
''    Class used to store/modify/restore Excel Application properties in order
''       to speed up code execution
'' Example usage:
''    Sub Test()
''        Dim app As New ExcelAppState: app.StoreState: app.Sleep
''        'Do Whatever
''        app.RestoreState
''    End Sub
''==============================================================================

Option Explicit

'Primary class members
Private m_enableEvents As Boolean
Private m_calculationMode As XlCalculation
Private m_screenUpdating As Boolean
Private m_displayAlerts As Boolean

'Keep track if a state is stored
Private m_hasStoredState As Boolean
'Keep track if Calculation Mode was stored
Private m_hasStoredCalcMode As Boolean

'*******************************************************************************
'Retain the current application settings
'*******************************************************************************
Public Sub StoreState()
    With Application
        m_enableEvents = .EnableEvents
        On Error Resume Next 'In case no Workbook is opened
        m_calculationMode = .Calculation
        m_hasStoredCalcMode = (Err.Number = 0)
        On Error GoTo 0
        m_screenUpdating = .ScreenUpdating
        m_displayAlerts = .DisplayAlerts
    End With
    m_hasStoredState = True
End Sub

'*******************************************************************************
'Restore the previously saved application settings
'*******************************************************************************
Public Sub RestoreState(Optional ByVal maxSecondsToWait As Integer)
    If Not m_hasStoredState Then
        Err.Raise 5, TypeName(Me) & ".RestoreState", "State not stored"
    End If
    With Application
        If .EnableEvents <> m_enableEvents Then .EnableEvents = m_enableEvents
        If m_hasStoredCalcMode Then
            On Error Resume Next
            If .Calculation <> m_calculationMode Then .Calculation = m_calculationMode
            If Err.Number = 0 Then WaitForCalculations maxSecondsToWait
            On Error GoTo 0
        End If
        If .ScreenUpdating <> m_screenUpdating Then .ScreenUpdating = m_screenUpdating
        If .DisplayAlerts <> m_displayAlerts Then .DisplayAlerts = m_displayAlerts
    End With
    m_hasStoredState = False
End Sub

'*******************************************************************************
'Put App to sleep. Speeds up code execution
'*******************************************************************************
Public Sub Sleep()
    With Application
        If .EnableEvents Then .EnableEvents = False
        On Error Resume Next
        If .Calculation <> xlCalculationManual Then .Calculation = xlCalculationManual
        On Error GoTo 0
        If .ScreenUpdating Then .ScreenUpdating = False
        If .DisplayAlerts Then .DisplayAlerts = False
    End With
End Sub

'*******************************************************************************
'Set Application to Normal working mode
'*******************************************************************************
Public Sub Wake(Optional ByVal maxSecondsToWait As Integer = 10)
    With Application
        If Not .EnableEvents Then .EnableEvents = True
        On Error Resume Next
        If .Calculation <> xlCalculationAutomatic Then .Calculation = xlCalculationAutomatic
        If Err.Number = 0 Then WaitForCalculations maxSecondsToWait
        On Error GoTo 0
        If Not .ScreenUpdating Then .ScreenUpdating = True
        If Not .DisplayAlerts Then .DisplayAlerts = True
    End With
End Sub

'*******************************************************************************
'Wait for the application to finish calculations
'*******************************************************************************
Public Sub WaitForCalculations(ByVal maxSecondsToWait As Integer)
    If Application.Calculation = xlCalculationManual Then Exit Sub
    If maxSecondsToWait <= 0 Then Exit Sub
    '
    Dim remainingTriesToFixBug As Long: remainingTriesToFixBug = 3
    Dim tStart As Date: tStart = VBA.Now()
    Dim tMax As Date: tMax = tStart + VBA.TimeSerial(0, 0, maxSecondsToWait)
    '
    Do While Application.CalculationState <> xlDone
        DoEvents
        If VBA.Now() > tMax Then Exit Do
        If Application.CalculationState = xlPending Then
            remainingTriesToFixBug = remainingTriesToFixBug - 1
            If remainingTriesToFixBug < 0 Then Exit Do
            TryFixingPendingBug
        End If
    Loop
End Sub

'*******************************************************************************
'Utility function for WaitForCalculations method
'*******************************************************************************
Private Sub TryFixingPendingBug()
    'There is a bug in Excel when a Volatile function is present: even if
    '   the Status Bar shows 'Ready', Application.CalculationState could
    '   still return 'xlPending'
    'Force state out of pending by marking any formula cell as needing
    '   recalculation
    Dim formulaCell As Range
    '
    Set formulaCell = GetEditableFormulaCell()
    If Not formulaCell Is Nothing Then formulaCell.Dirty
End Sub

'*******************************************************************************
'Utility function for TryFixingPendingBug method
'*******************************************************************************
Private Function GetEditableFormulaCell() As Range
    Dim wSheet As Worksheet
    Dim formulaCell As Range
    '
    For Each wSheet In ThisWorkbook.Worksheets
        Set formulaCell = GetEditableFormulaCellFromSheet(wSheet)
        If Not formulaCell Is Nothing Then Exit For
    Next wSheet
    Set GetEditableFormulaCell = formulaCell
End Function

'*******************************************************************************
'Utility function for GetEditableFormulaCell method
'*******************************************************************************
Private Function GetEditableFormulaCellFromSheet(wSheet As Worksheet) As Range
    If wSheet.EnableCalculation Then
        Dim rngFormulas As Range
        '
        On Error Resume Next
        Set rngFormulas = wSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If rngFormulas Is Nothing Then Exit Function
        '
        'Check if Worksheet is Macro Protected
        If (wSheet.ProtectContents Or wSheet.ProtectDrawingObjects _
        Or wSheet.ProtectScenarios) And Not wSheet.ProtectionMode _
        Then
            Dim rngCell As Range
            '
            For Each rngCell In rngFormulas
                If Not rngCell.Locked Then
                    Set GetEditableFormulaCellFromSheet = rngCell
                    Exit Function
                End If
            Next rngCell
        Else
            Set GetEditableFormulaCellFromSheet = rngFormulas.Cells(1, 1)
        End If
    End If
End Function

