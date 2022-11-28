Attribute VB_Name = "LibUDFs"
'''=============================================================================
''' VBA FastExcelUDFs
'''--------------------------------------------------
''' https://github.com/cristianbuse/VBA-FastExcelUDFs
'''--------------------------------------------------
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
''  - Having a large number of User Defined Functions (UDFs) can be very slow
''    because of a known bug in Excel that causes the state of the VBE window to
''    be updated for each UDF
''  - Note that the bug is not present if the calculation is triggered from
''    outside of the UDF context e.g. from a macro
''  - If VBE is open and UDFs are calculating then the VBE is flickering and a
''    word [Running] can be observed in the caption
'' Solution:
''  - A call to 'TriggerFastUDFCalculation' method must be placed in all UDFs
''    so that an async call can do the calculation outside of the UDF context
''  - The async call is done via the 'QueryClose' event of a form by posting
''    a WM_DESTROY message to the form's window
''  - For stability, no more API Timers are used
''  - A Mouse Input (a mild horizontal scroll) is sent to the Application
''    using the 'SendInput' API - to get Excel out of Calculation mode
''  - Once Excel is out of Calculation mode the async call will trigger a
''    calculation outside of the UDF context thus avoiding the mentioned bug
'' Notes:
''  - This solution works only if the 'CalculationInterruptKey' property of the
''    Excel Application is set to 'xlAnyKey' (default)
''  - Flags are in place to minimize the number of calculations needed
''==============================================================================

Option Explicit
Option Private Module

#If Mac Then
#ElseIf VBA7 Then
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendInput Lib "user32" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByVal hWnd As Long) As Long
#End If

'Necessary Structures for SendInput API
'Note that GENERALINPUT is simplified to work with MOUSEINPUT (ignored Keyboard)
'https://docs.microsoft.com/en-gb/windows/desktop/api/winuser/ns-winuser-taginput
Private Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    #If Win64 Then
        dwExtraInfo As LongLong
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

Private m_calculationInProgress As Boolean
Private m_asyncForm As AsyncFormCall
Private m_lastCaller As Range

'*******************************************************************************
'Try to trigger a calculation outside of the UDF context
'*******************************************************************************
Public Sub TriggerFastUDFCalculation()
    If m_calculationInProgress And IsFormConnected(m_asyncForm) Then Exit Sub
    '
    If ThisWorkbook.IsAddin Then Exit Sub
    If Application.Calculation = xlCalculationManual Then Exit Sub
    If Application.CalculationInterruptKey <> xlAnyKey Then Exit Sub
    '
    m_calculationInProgress = True
    Set LastCallerRange = Application.Caller
    MakeAsyncCall
    InterruptCalculation
End Sub
Private Property Set LastCallerRange(ByRef rCaller As Variant)
    If TypeName(rCaller) = "Range" Then
        Set m_lastCaller = rCaller
    Else
        Set m_lastCaller = Nothing
    End If
End Property

'*******************************************************************************
'Generate an async callback outside of the UDF context
'*******************************************************************************
Private Function MakeAsyncCall()
    Const WM_DESTROY As Long = &H2
    Static hWnd As LongPtr
    '
    If Not IsFormConnected(m_asyncForm) Then
        Set m_asyncForm = New AsyncFormCall
        #If Mac = 0 Then
            IUnknown_GetWindow m_asyncForm, VBA.VarPtr(hWnd)
        #End If
    End If
    m_asyncForm.EnableCall = True
    #If Mac = 0 Then
        PostMessage hWnd, WM_DESTROY, 0, 0
    #End If
End Function
Private Function IsFormConnected(ByVal obj As Object) As Boolean
    If Not obj Is Nothing Then
        IsFormConnected = TypeName(obj) <> "UserForm"
    End If
End Function

'*******************************************************************************
'Generate a fake User Interaction Event in order to pause calculation
'*******************************************************************************
Private Sub InterruptCalculation()
    Const INPUT_MOUSE As Long = 0&
    Const MOUSEEVENTF_HWHEEL = &H1000 'Horizontal Wheel Scroll
    Dim GInput As GENERALINPUT
    Static s As Long
    '
    If s = 0 Then s = 1 Else s = -s
    '
    GInput.dwType = INPUT_MOUSE
    GInput.mi.dwFlags = MOUSEEVENTF_HWHEEL
    GInput.mi.mouseData = s 'Must be different from 0 for the interrupt to work
    '
    #If Mac = 0 Then
        SendInput 1, GInput, Len(GInput)
    #End If
End Sub

'*******************************************************************************
'Calculate UDFs
'*******************************************************************************
Public Sub FastCalculate()
    m_calculationInProgress = True
    If IsFormConnected(m_asyncForm) Then m_asyncForm.EnableCall = False
    '
    Dim cKey As XlEnableCancelKey: cKey = Application.EnableCancelKey
    Dim tries As Long: tries = 3
    '
    On Error Resume Next
    Application.Cursor = xlWait
    Application.EnableCancelKey = xlDisabled
    Do While Application.CalculationState <> xlDone
        DoEvents
        If Application.CalculationState = xlPending Then
            If m_lastCaller Is Nothing Then Exit Do
            If tries = 0 Then Exit Do
            '
            m_lastCaller.Dirty
            tries = tries - 1
        End If
    Loop
    Application.Cursor = xlDefault
    Application.EnableCancelKey = cKey
    On Error GoTo 0
    '
    m_calculationInProgress = False
End Sub
