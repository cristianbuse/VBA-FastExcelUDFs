# VBA-FastExcelUDFs

FastExcelUDFs is a VBA Project that allows faster User Defined Function (UDF) calculation when Excel is in Automatic Calculation mode.

Having a large number of User Defined Functions (UDFs) can be very slow because of a known bug in Excel that causes the state of the VBE window to be updated for each UDF called when in Automatic Calculation mode. Note that the bug is not present in Manual Calculation mode. If VBE is opened and UDFs are calculating the VBE is flickering and a word [Running] can be observed in the caption.

**Solution**

A call to ```TriggerFastUDFCalculation``` method must be placed in all UDFs.
The first time this method is called (per each calculation session), 3 things are happening:
 1. a boolean flag is set (m_fastOn) in order to run the logic only once per calculation session
 2. a Timer is set using Windows API - a callback will be triggered as soon as Excel gets out of Calculation mode
 3. a Mouse Input (a mild horizontal scroll) is sent to the Application using a Windows API - this is done in order to get Excel out of Calculation mode. Once Excel is out of Calculation mode the Timer will kick in and the Application will calculate in Manual mode to avoid the mentioned bug.

**Newer Excel versions**

Although it might seem that the bug is not present (i.e. there is no visible difference between using and not using the solution presented here), the bug is still manifesting if the VBA IDE is opened via Alt+F11 or via Developer ribbon tab/Visual Basic.


## Installation

Just import the following 2 code modules in your VBA Project:

* **LibUDFs.bas**  
* **ExcelAppState.cls**

## Usage
In your UDFs use:
```vba
LibUDFs.TriggerFastUDFCalculation
```
For example:
```vba
Public Function TEST_UDF(value As Boolean) As Variant
    Application.Volatile False
    LibUDFs.TriggerFastUDFCalculation
    '
    TEST_UDF = value
End Function
```

## Notes
* The Fast Calculation is only triggered if a Calculation Interruption is possible. For that make sure:
```vba
Application.CalculationInterruptKey = xlAnyKey
```
* You can download the available Demo Workbook for a quick start

## License
MIT License

Copyright (c) 2020 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.