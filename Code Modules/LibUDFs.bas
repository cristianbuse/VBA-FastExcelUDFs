Attribute VB_Name = "LibUDFs"
'''=============================================================================
'' VBA FastExcelUDFs
'''-------------------------------------------------
'' https://github.com/cristianbuse/VBA-FastExcelUDFs
'''-------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2019 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================
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
Option Private Module

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
    On Error GoTo ErrorHandler
    If Application.Calculation = xlCalculationManual Then Exit Sub
    On Error GoTo 0
    If Application.CalculationInterruptKey = xlAnyKey Then
        m_fastOn = True
        StartTimer milliSeconds:=10
        ForceCalculationInterruption
    Else
        '[Application.CalculationInterruptKey] must be set to xlAnyKey in order
        '   to trigger FastUDF calculation
    End If
Exit Sub
ErrorHandler:
    'Calling this function from an .xlam AddIn when no workbooks are opened
    '   would generate an error on reading the Application.Calculation property
End Sub

'*******************************************************************************
'Generate a fake User Interaction Event in order to pause calculation
'*******************************************************************************
Private Sub ForceCalculationInterruption()
    Const INPUT_MOUSE As Long = 0&
    Const MOUSEEVENTF_HWHEEL = &H1000 'Horizontal Wheel Scroll
    Dim GInput As GENERALINPUT
    '
    GInput.dwType = INPUT_MOUSE
    GInput.mi.dwFlags = MOUSEEVENTF_HWHEEL
    GInput.mi.mouseData = 1 'Must be different from 0
    '
    #If Mac Then
    #Else
        SetFocus Application.hwnd
        SendInput 1, GInput, Len(GInput)
    #End If
End Sub

'*******************************************************************************
'Sets a timer that will call back 'TimerProc' outside of the UDF context (this
'   allows VBA code to alter Application State)
'*******************************************************************************
Private Sub StartTimer(ByVal milliSeconds As Long)
    #If Mac Then
    #Else
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
    #If Mac Then
    #Else
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
        Application.ScreenUpdating = True
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
