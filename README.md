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
Copyright (C) 2019 VBA Mouse Scroll project contributors

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/) or
[GPLv3](https://choosealicense.com/licenses/gpl-3.0/).