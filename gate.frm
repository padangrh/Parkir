VERSION 5.00
Begin VB.Form Form_Gate 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buka Gate"
   ClientHeight    =   3030
   ClientLeft      =   6540
   ClientTop       =   4380
   ClientWidth     =   7230
   Icon            =   "gate.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cb_Keterangan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "gate.frx":0CCA
      Left            =   1800
      List            =   "gate.frx":0CE0
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton btn_Cancel 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton btn_Save 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txt_Keterangan 
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
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gunakan tombol ini hanya ketika pintu tidak terbuka atau terjadi kesalahan. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   360
      Picture         =   "gate.frx":0D60
      Stretch         =   -1  'True
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "Form_Gate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_Cancel_Click()
    Unload Me
End Sub

Private Sub btn_Save_Click()
    Dim filename As String
    filename = Left(Form_Main.kode_Transaksi, 1) & Format(Now, "yyyy-MM-dd-h-mm-ss") & ".jpg"
    Call Form_Main.take_pic("\Simpan\", filename)
    'Form_Main.tutup_foto
    con.Execute ("insert into tbbuka values('" & Form_Main.kode_Transaksi & "', '" & Format(Now, "yyyy-MM-dd") & "', '" & Format(Now, "HH:mm:ss") & "', '" & username & "','" & cb_Keterangan.Text & " - " & txt_Keterangan & "', '" & filename & "')")
    Form_Main.nextFaktur
    Form_Main.Buka_Pintu
    Unload Me
End Sub

Private Sub Form_Load()
    cb_Keterangan.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Main.Timer1.Enabled = True
End Sub

Private Sub txt_Keterangan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        
        'MainForm.simpanFotoGate
        'MainForm.bukaGate
        btn_Save_Click
        'Unload Me
    End If
End Sub
