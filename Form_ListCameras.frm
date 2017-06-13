VERSION 5.00
Begin VB.Form Form_ListCameras 
   Caption         =   "Daftar Kamera"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3660
      TabIndex        =   2
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   4920
      Width           =   855
   End
   Begin VB.ListBox lstFilters 
      Height          =   4350
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4275
   End
End
Attribute VB_Name = "Form_ListCameras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Oked As Boolean
Public CameraName As String

Private Sub cmdCancel_Click()
    Oked = False
    lstFilters.SetFocus
    Hide
End Sub

Private Sub cmdOk_Click()
    Oked = True
    With lstFilters
        CameraName = .List(.ListIndex)
    End With
    lstFilters.SetFocus
    Hide
End Sub

Private Sub Form_Activate()
    Dim rfiEach As QuartzTypeLib.IRegFilterInfo
    Dim x As Integer
    lstFilters.Clear
    With New QuartzTypeLib.FilgraphManager
        For Each rfiEach In .RegFilterCollection
            'x = -1
            'x = InStr(rfiEach.Name, "cam")
            'If x > 0 Then
            lstFilters.AddItem rfiEach.Name
            
        Next
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Oked = False
        Hide
    End If
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        lstFilters.Width = ScaleWidth
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Main.Timer1.Enabled = True
End Sub

Private Sub lstFilters_Click()
    cmdOk.Enabled = lstFilters.ListIndex > -1
End Sub


