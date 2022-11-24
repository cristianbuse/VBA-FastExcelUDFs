Attribute VB_Name = "DEMO"
Option Explicit

Public Function TEST_UDF(value As Boolean) As Variant
    Application.Volatile False
    LibUDFs.TriggerFastUDFCalculation
    '
    TEST_UDF = value
End Function

Public Function TEST_UDF2(value As Boolean) As Variant
    Application.Volatile False
    '
    TEST_UDF2 = value
End Function
