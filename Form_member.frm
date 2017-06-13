VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form_Member 
   Caption         =   "Member Manager"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Label 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox txt_Perusahaan 
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
      Left            =   4920
      TabIndex        =   3
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txt_Phone 
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
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
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
      ItemData        =   "Form_member.frx":0000
      Left            =   4920
      List            =   "Form_member.frx":000A
      TabIndex        =   4
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txt_nama 
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
      Left            =   4920
      TabIndex        =   1
      Top             =   2640
      Width           =   2775
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
      Left            =   4920
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
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
      Height          =   4740
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
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
      Left            =   5520
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
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
      Left            =   4440
      Top             =   6120
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
            Picture         =   "Form_member.frx":0020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_member.frx":0DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_member.frx":1DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_member.frx":2D10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Perusahaan"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   3240
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
      Left            =   3240
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RFID"
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
      Left            =   3240
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Member Manager"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "form_Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMember As ADODB.Recordset

Private Sub Command1_Click()
    reset
End Sub

Private Sub Form_Load()
    reload
    cb_status.ListIndex = 0
End Sub

Private Sub reload()
    list_user.Clear
    Set rsMember = con.Execute("select * from tbmember")
    If rsMember.EOF Then Exit Sub
    
    rsMember.MoveFirst
    Do While Not rsMember.EOF
        list_user.AddItem (rsMember!memberid)
        rsMember.MoveNext
    Loop
End Sub

Private Function getmember(memberid As String) As Boolean
    Dim found As Boolean
    found = False
    rsMember.MoveFirst
    Do While Not rsMember.EOF
        If rsMember!memberid = memberid Then
            found = True
            Exit Do
        End If
        rsMember.MoveNext
    Loop
    
    getmember = found
End Function

Private Sub Form_Unload(Cancel As Integer)
    Form_Main.Timer1.Enabled = True
End Sub

Private Sub list_user_Click()
    If getmember(list_user.Text) Then
        txt_id = rsMember!memberid
        txt_nama = rsMember!nama
        txt_Phone = rsMember!phone
        txt_Perusahaan = rsMember!perusahaan
        If rsMember!status = 1 Then
            cb_status.Text = "Aktif"
        Else
            cb_status.Text = "Non-Aktif"
        End If
        txt_Label = rsMember!Label
    Else
        reset
    End If
End Sub

Private Sub reset()
    txt_id = ""
    txt_nama = ""
    txt_Phone = ""
    txt_Perusahaan = ""
    cb_status.ListIndex = 0
    txt_Label = ""
    txt_id.SetFocus
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
    Dim temp_status As Integer
    If txt_id = "" Or Len(txt_id) <> 10 Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If getmember(txt_id) Then
        MsgBox "RFID telah terpakai"
        Exit Sub
    End If
    
    If cb_status.Text = "Non-Aktif" Then
        temp_status = 0
    Else
        temp_status = 1
    End If
    
    con.Execute ("insert into tbmember values('" & txt_id & "', '" & txt_nama & "', '" & txt_Phone & "', '" & txt_Perusahaan & "' , " & Val(temp_status) & ", '" & txt_Label & "')")
    reload
    reset
End Sub

Private Sub ubah()
    Dim temp_status As Integer
    
    If txt_id = "" Or Len(txt_id) <> 10 Then
        MsgBox "Data tidak lengkap"
        Exit Sub
    End If
    
    If Not getmember(txt_id) Then
        MsgBox "RFID tidak ditemukan"
        Exit Sub
    End If
    
    If cb_status.Text = "Non-Aktif" Then
        temp_status = 0
    Else
        temp_status = 1
    End If
    
    con.Execute ("update tbmember set nama = '" & txt_nama & "', phone = '" & txt_Phone & "', perusahaan = '" & txt_Perusahaan & "' , status = " & Val(temp_status) & ", label = '" & txt_Label & "' where memberid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub hapus()
    con.Execute ("delete from tbmember where memberid = '" & txt_id & "'")
    reload
    reset
End Sub

Private Sub keluar()
    Unload Me
End Sub

Private Sub txt_id_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Len(txt_id) = 10 Then
        If getmember(txt_id) Then
            txt_id = rsMember!memberid
            txt_nama = rsMember!nama
            txt_Phone = rsMember!phone
            txt_Perusahaan = rsMember!perusahaan
            If rsMember!status = 1 Then
                cb_status.Text = "Aktif"
            Else
                cb_status.Text = "Non-Aktif"
            End If
            txt_Label = rsMember!Label
        Else
            reset
        End If
    End If
End Sub
