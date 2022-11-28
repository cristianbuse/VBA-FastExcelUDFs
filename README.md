# VBA-FastExcelUDFs

FastExcelUDFs is a VBA Project that allows faster User Defined Function (UDF) calculation.

Having a large number of User Defined Functions (UDFs) can be very slow because of a [known bug in Excel](https://fastexcel.wordpress.com/2011/06/13/writing-efficient-vba-udfs-part-3-avoiding-the-vbe-refresh-bug/) that causes the state of the VBE window to be updated for each UDF called. Note that the bug is not present outside of the UDF context e.g. when calculation is triggered from a macro. If VBE is open and UDFs are calculating then the VBE is flickering and a word [Running] can be observed in the caption.

**Solution**

A call to ```TriggerFastUDFCalculation``` method must be placed in all UDFs so that an async call can do the calculation outside of the UDF context.
- The async call is done via the ```QueryClose``` event of a form by posting a ```WM_DESTROY``` message to the form's window. For stability, no more API Timers are used (see [old version](https://github.com/cristianbuse/VBA-FastExcelUDFs/tree/a196f1bf830d4e9e6fb0a14cdd81462bffcc0433) using timers - causes crashes particularly on x64).
- A Mouse Input (a mild horizontal scroll) is sent to the Application using the 'SendInput' API - to get Excel out of Calculation mode.
- Once Excel is out of Calculation mode the async call will trigger a calculation outside of the UDF context thus avoiding the mentioned bug.

**Newer Excel versions**

Although it might seem that the bug is not present (i.e. there is no visible difference between using and not using the solution presented here), the bug is still manifesting if the VBA IDE is opened via Alt+F11 or via Developer ribbon tab/Visual Basic.


## Installation

Just import the following 2 code modules in your VBA Project. To avoid copy-paste issues please download [zip](https://github.com/cristianbuse/VBA-FastExcelUDFs/archive/refs/heads/master.zip) and then import from there.

* **LibUDFs.bas**  
* **AsyncFormCall.frm**

If importing the form then please note that the *AsyncFormCall.frx* file is also needed when importing. Alternatively, you can recreate the form in 3 simple steps:
1) insert a new userforn
2) rename it to ```AsyncFormCall```
3) Paste the following code in the form code module:  
```VBA
Option Explicit

Private m_enableCall As Boolean

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    If m_enableCall Then FastCalculate
End Sub

Public Property Let EnableCall(ByVal newValue As Boolean)
    m_enableCall = newValue
End Property
```

## Usage
In your UDFs use:  
```VBA
LibUDFs.TriggerFastUDFCalculation
```
For example:  
```VBA
Public Function TEST_UDF(ByVal value As Boolean) As Variant
    Application.Volatile False
    LibUDFs.TriggerFastUDFCalculation
    '
    TEST_UDF = value
End Function
```

## Notes
* The Fast Calculation is only triggered if a Calculation Interruption is possible. For that make sure:  
```VBA
Application.CalculationInterruptKey = xlAnyKey
```
* You can download the available [Demo Workbook](https://github.com/cristianbuse/VBA-FastExcelUDFs/blob/master/VBA%20FastExcelUDFs_DEMO.xlsm) for a quick start

## License
MIT License

Copyright (c) 2019 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.