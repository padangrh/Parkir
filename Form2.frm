VERSION 5.00
Begin VB.Form Form_Login 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Login"
   ClientHeight    =   3105
   ClientLeft      =   7125
   ClientTop       =   4455
   ClientWidth     =   6090
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   6090
   Begin VB.TextBox txtuser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Commandlogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Commandbatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   5760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandBatal_Click()
End
End Sub

Private Sub CommandLogin_Click()

On Error GoTo salah:
    Dim Rec As ADODB.Recordset
    Set Rec = con.Execute("select * from tblogin where userid='" & Trim(txtuser.Text) & "'")
    If Not Rec.EOF Then
        If UCase(Rec.Fields("userid")) = UCase(Trim(txtuser)) And Rec.Fields("pass") = Trim(txtpass) Then
            'username = Rec!username
            'status = Rec!posisi
            
            'FrmMain.p.Enabled = CBool(Rec.Fields("hak1"))
            'FrmMain.Toolbar1.Buttons(1).Enabled = CBool(Rec.Fields("hak1"))
            'FrmMain.l.Enabled = CBool(Rec.Fields("hak2"))
            'FrmMain.Toolbar1.Buttons(2).Enabled = CBool(Rec.Fields("hak2"))
            'FrmMain.b.Enabled = CBool(Rec.Fields("hak3"))
            'FrmMain.Toolbar1.Buttons(3).Enabled = CBool(Rec.Fields("hak3"))
            'FrmMain.a.Enabled = CBool(Rec.Fields("hak4"))
            'FrmMain.Toolbar1.Buttons(4).Enabled = CBool(Rec.Fields("hak4"))
            username = Rec.Fields("userid")
            status = Rec.Fields("posisi")
            Unload Me
            Form_Main.Show (1)
            
            'Form_Main.Toolbar1.Enabled = True
        Else
            MsgBox "Nama user atau password anda tidak cocok!"
            txtuser.SetFocus
        End If
    Else
      MsgBox "Nama user atau password anda tidak cocok!"
      txtuser.SetFocus
    End If
    Exit Sub
    
salah:
MsgBox "Periksa komputer server hidup atau tidak, kabel internet tercolok di komputer atau tidak, coba restart modem"
End Sub

Private Sub Form_Activate()
    connect
    txtuser.SetFocus
End Sub

Private Sub txtpass_keyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        CommandLogin_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'con.Close
End Sub



