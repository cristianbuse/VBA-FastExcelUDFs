# VBA-FastExcelUDFs

FastExcelUDFs is a VBA Project that allows faster User Defined Function (UDF) calculation when Excel is in Automatic Calculation mode.

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