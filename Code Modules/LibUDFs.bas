Attribute VB_Name = "LibUDFs"
''==============================================================================
'' VBA FastExcelUDFs
''---------------------------------
'' Copyright (C) 2019 Cristian Buse
'' https://github.com/cristianbuse/VBA-FastExcelUDFs
''---------------------------------
'' This program is free software: you can redistribute it and/or modify
'' it under the terms of the GNU General Public License as published by
'' the Free Software Foundation, either version 3 of the License, or
'' (at your option) any later version.
''
'' This program is distributed in the hope that it will be useful,
'' but WITHOUT ANY WARRANTY; without even the implied warranty of
'' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'' GNU General Public License for more details.
''
'' You should have received a copy of the GNU General Public License
'' along with this program.  If not, see <https://www.gnu.org/licenses/>.
''==============================================================================
''
''==============================================================================
'' Description:
''    Having a large number of User Defined Functions (UDFs) can be very slow
''       because of a known bug in Excel that causes the state of the VBE window
''       to be updated for each UDF called when in Automatic Calculation mode.
''    Note that the bug is not present in Manual Calculation mode.
''    If VBE is opened and UDFs are calculating the VBE is flickering and a
''        word [Running] can be observed in the caption.
'' Solution:
''    A call to <TriggerFastUDFCalculation> method must be placed in all UDFs.
''    The first time this method is called, 3 things happen:
''        1) a boolean flag is set (m_fastOn) in order to run the logic only
''           once per calculation session
''        2) a Timer is set using Windows API - a callback will be triggered as
''           soon as Excel gets out of Calculation mode
''        3) A Mouse Input (a mild horizontal scroll) is sent to the Application
''           using a Windows API - this is done in order to get Excel out of
''           Calculation mode
''    Once Excel is out of Calculation mode the Timer will kick in and the
''        Application will calculate in Manual mode to avoid the mentioned bug.
'' Notes:
''    The above mentioned solution works only if the .CalculationInterruptKey
''       property of the Excel Application is set to 'xlAnyKey' (default)
''    The Timer terminates itself after the first call to minimize the chance of
''       generating a crash (see warning below)
'' Warning:
''    Do not debug code while the Timer's callback has not been called. This
''      will cause a crash particularly on x64 versions of Excel
'' Requires:
''    - ExcelAppState: class that can store/modify/restore Excel App properties
''==============================================================================

Option Explicit

'Windows APIs
'*******************************************************************************
#If Mac Then
    'Support not available
#ElseIf VBA7 Then
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private Declare PtrSafe Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
    Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
#Else
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
    Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
#End If
'*******************************************************************************

'Necessary Structures for SendInput API
'Note that GENERALINPUT is simplified to work with MOUSEINPUT (ignored Keyboad)
'https://docs.microsoft.com/en-gb/windows/desktop/api/winuser/ns-winuser-taginput
'*******************************************************************************
Private Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    #If Win64 Then
        dwExtraInfo As LongPtr
        dummyMemoryOffset As Long
    #Else
        dwExtraInfo As Long
    #End If
End Type
Private Type GENERALINPUT
    dwType As Long
    #If Win64 Then
        dummyMemoryOffset As Long
    #End If
    mi As MOUSEINPUT
End Type
'*******************************************************************************

'Boolean tracking if TriggerFastUDFCalculation was already called
Private m_fastOn As Boolean

'*******************************************************************************
'Prepare environment to calculate UDFs in Manual Calculation Mode
'*******************************************************************************
Public Sub TriggerFastUDFCalculation()
    If m_fastOn Then Exit Sub
    m_fastOn = True
    StartTimer milliSeconds:=10
    ForceCalculationInterruption
End Sub

'*******************************************************************************
'Generate a fake User Interaction Event in order to pause calculation
'*******************************************************************************
Private Sub ForceCalculationInterruption()
    If Application.CalculationInterruptKey = xlAnyKey Then
        Const INPUT_MOUSE As Long = 0&
        Const MOUSEEVENTF_HWHEEL = &H1000 'Horizontal Wheel Scroll
        Dim GInput As GENERALINPUT
        '
        GInput.dwType = INPUT_MOUSE
        GInput.mi.dwFlags = MOUSEEVENTF_HWHEEL
        GInput.mi.mouseData = 1 'Must be different from 0
        '
        #If Not Mac Then
            SetFocus Application.hwnd
            SendInput 1, GInput, Len(GInput)
        #End If
    Else
        Debug.Print "[Application.CalculationInterruptKey] must be set to " _
            & "'xlAnyKey' in order to trigger FastUDF calculation"
    End If
End Sub

'*******************************************************************************
'Sets a timer that will call back 'TimerProc' outside of the UDF context (this
'   allows VBA code to alter Application State)
'*******************************************************************************
Private Sub StartTimer(ByVal milliSeconds As Long)
    #If Not Mac Then
        SetTimer Application.hwnd, 0&, ByVal milliSeconds, AddressOf TimerProc
    #End If
End Sub

'*******************************************************************************
'The Timer Callback Function.
'The Timer is 'killed' immediately after being triggered to make sure it only
'   runs once per UDF trigger
'*******************************************************************************
#If VBA7 Then
Private Sub TimerProc(ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal nIDEvent As LongPtr, ByVal wTime As Long)
#Else
Private Sub TimerProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal nIDEvent As Long, ByVal wTime As Long)
#End If
    #If Not Mac Then
        If KillTimer(hwnd, nIDEvent) Then FastCalculate
    #End If
End Sub

'*******************************************************************************
'Calculate UDFs in Manual Calculation Mode
'*******************************************************************************
Private Sub FastCalculate()
    If m_fastOn Then
        On Error GoTo ErrorHandler
        Dim app As New ExcelAppState: app.StoreState: app.Sleep
        Application.Calculate
        app.RestoreState
        m_fastOn = False
    End If
Exit Sub
ErrorHandler:
    'Application State cannot be modified
    'Most likely a UDF will restart the whole triggering process
    m_fastOn = False
End Sub
