VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AsyncFormCall 
   Caption         =   "Async"
   ClientHeight    =   615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "AsyncFormCall.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AsyncFormCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_enableCall As Boolean

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = True
    If m_enableCall Then FastCalculate
End Sub

Public Property Let EnableCall(ByVal newValue As Boolean)
    m_enableCall = newValue
End Property
