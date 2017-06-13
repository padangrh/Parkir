VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_User 
   Caption         =   "User Manager"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ListBox list_user 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txt_id 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txt_password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.ComboBox cb_status 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form_User.frx":0000
      Left            =   4080
      List            =   "Form_User.frx":000D
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   1614
      ButtonWidth     =   3043
      ButtonHeight    =   1455
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":002F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":0DC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":1DF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_User.frx":2D1F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "Form_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As ADODB.Recordset

Private Sub Command1_Click()
    reset
End Sub

Private Sub Form_Load()
    reload
End Sub

Private Sub reload()
    list_user.Clear
    Set rsUser = con.Execute("select * from tblogin")
    If rsUser.EOF Then Exit Sub
    
    rsUser.MoveFirst
    Do While Not rsUser.EOF
        list_user.AddItem (rsUser!userid)
        rsUser.MoveNext
    Loop
End Sub

Private Function getUser(userid As String) As Boolean
    Dim found As Boolean
    found = False
    rsUser.MoveFirst
    Do While Not rsUser.EOF
        If rsUser!userid = userid Then
            found = True
            Exit Do
        End If
        rsUser.MoveNext
    Loop
    
    getUser = found
End Function

Private Sub Form_Unload(Cancel As Integer)
    Form_Main.Timer1.Enabled = True
End Sub



Private Sub list_user_Click()
    If getUser(list_user.Text) Then
        txt_id = rsUser!userid
        txt_password = rsUser!pass
        cb_status.Text = rsUser!posisi
    Else
        reset
    End If
End Sub

Private Sub reset()
    txt_id = ""
    txt_password = ""
    cb_status.ListIndex = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: tambah
        Case 2: ubah
        Case 3: hapus
        Case 4: keluar
    End Select
End Sub

Private Sub tambah()
    If txt_id = "" Or txt_password = "" Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If getUser(txt_id) Then
        MsgBox "User ID telah terpakai"
        Exit Sub
    End If
    
    con.Execute ("insert into tblogin values('" & txt_id & "', '" & txt_password & "', '" & cb_status.Text & "')")
    reload
    reset
End Sub

Private Sub ubah()
    If txt_id = "" Or txt_password = "" Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If Not getUser(txt_id) Then
        MsgBox "User ID tidak ditemukan"
        Exit Sub
    End If
    
    con.Execute ("update tblogin set pass = '" & txt_password & "', posisi = '" & cb_status.Text & "' where userid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub hapus()
    con.Execute ("delete from tblogin where userid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub keluar()
    Unload Me
End Sub

