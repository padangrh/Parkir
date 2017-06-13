VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Laporan 
   Caption         =   "Laporan"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   6975
      Begin MSComctlLib.ListView lv_memberId 
         Height          =   1215
         Left            =   2640
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MemberID"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nama"
            Object.Width           =   2822
         EndProperty
      End
      Begin VB.CommandButton btn_memberPerorangan 
         Caption         =   "Laporan Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         TabIndex        =   17
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txt_memberID 
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
         Left            =   2640
         TabIndex        =   16
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Member ID"
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
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton btn_Review 
      Caption         =   "Review"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   13
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton btn_LaporanMember 
      Caption         =   "Laporan Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CheckBox chk_Sampai 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Perorangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6975
      Begin MSComctlLib.ListView lv_Userid 
         Height          =   1215
         Left            =   2640
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Username"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Posisi"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txt_Userid 
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
         Left            =   2640
         TabIndex        =   5
         Top             =   480
         Width           =   3015
      End
      Begin VB.CommandButton btn_HarianPersonal 
         Caption         =   "Laporan Parkir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label3 
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton btn_Darurat 
      Caption         =   "Laporan Darurat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton btn_LaporanParkir 
      Caption         =   "Laporan Parkir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin Crystal.CrystalReport cr 
      Left            =   7080
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dt_start 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97255425
      CurrentDate     =   42810
   End
   Begin MSComCtl2.DTPicker dt_end 
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97255425
      CurrentDate     =   42810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sampai"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form_Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Darurat_Click()
    
    If chk_Sampai.Value = 0 Then
        Call openReport("LaporanPintuDarurat.rpt", "tbbuka.tanggal", "ONE_DAY", False)
    Else
        Call openReport("LaporanPintuDarurat.rpt", "tbbuka.tanggal", "DURATION", False)
    End If
    
End Sub

Private Sub btn_HarianPersonal_Click()
    If check_userID(txt_Userid) = True Then
        If chk_Sampai.Value = 0 Then
            Call openReport("laporanharian3.rpt", "laporan.tanggal", "ONE_DAY", True)
        Else
            Call openReport("laporanharian4.rpt", "laporan.tanggal", "DURATION", True)
        End If
    Else
        MsgBox "User tidak ditemukan."
    End If
    
End Sub

Private Sub btn_LaporanMember_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("laporanmember.rpt", "tbtransaksi.tanggal", "ONE_DAY", False)
    Else
        Call openReport("laporanmember.rpt", "tbtransaksi.tanggal", "DURATION", False)
    End If
End Sub

Private Sub btn_LaporanParkir_Click()
    If chk_Sampai.Value = 0 Then
        Call openReport("LaporanParkir.rpt", "laporan.tanggal", "ONE_DAY", False)
    Else
        Call openReport("LaporanParkir.rpt", "laporan.tanggal", "DURATION", False)
    End If
End Sub

Private Sub btn_memberPerorangan_Click()
    If check_memberID(txt_memberID) = True Then
        If chk_Sampai.Value = 0 Then
            Call openReport("laporanmember.rpt", "tbtransaksi.tanggal", "ONE_DAY", True)
        Else
            Call openReport("laporanmember.rpt", "tbtransaksi.tanggal", "DURATION", True)
        End If
    Else
        MsgBox "Member tidak ditemukan."
    End If
End Sub

Private Sub btn_Review_Click()
    Form_Review.Show vbModal, Me
End Sub

Private Sub chk_Sampai_Click()
    If chk_Sampai.Value = 0 Then
        dt_end.Enabled = False
    Else
        dt_end.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    dt_end.Enabled = False
    dt_start.Value = Now
    dt_end.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Main.Timer1.Enabled = True
End Sub

Private Sub lv_memberId_DblClick()
    txt_memberID = lv_memberId.SelectedItem.Text
    btn_memberPerorangan.SetFocus
End Sub

Private Sub lv_memberId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call lv_memberId_DblClick
    End If
End Sub

Private Sub lv_memberId_LostFocus()
    lv_memberId.Visible = False
End Sub

Private Sub lv_Userid_DblClick()
    txt_Userid = lv_Userid.SelectedItem.Text
    btn_HarianPersonal.SetFocus
End Sub

Private Sub lv_Userid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
         Call lv_Userid_DblClick
    End If
End Sub

Private Sub lv_Userid_LostFocus()
    lv_Userid.Visible = False
End Sub

Private Sub txt_memberID_Change()
    If txt_memberID.Text <> "" Then
        lv_memberId.Visible = True
        reload_Member
    Else
        lv_memberId.Visible = False
    End If
End Sub

Private Sub txt_memberID_KeyDown(Key As Integer, Shift As Integer)
    If Key = 13 And lv_memberId.Visible = True Then
        lv_memberId.SetFocus
    ElseIf Key = 13 And lv_memberId.Visible = False Then
        btn_memberPerorangan.SetFocus
    End If
End Sub

Private Sub txt_Userid_Change()
    If txt_Userid.Text <> "" Then
        lv_Userid.Visible = True
        reload_List
    Else
        lv_Userid.Visible = False
    End If
End Sub


Private Sub txt_Userid_KeyDown(Key As Integer, Shift As Integer)
    If Key = 13 And lv_Userid.Visible = True Then
        lv_Userid.SetFocus
    ElseIf Key = 13 And lv_Userid.Visible = False Then
        btn_HarianPersonal.SetFocus
    End If
End Sub

Public Sub reload_List()
    lv_Userid.ListItems.Clear
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("select * from tblogin where userid like '%" & txt_Userid.Text & "%'")
    
    If rsFilter.EOF Then
        lv_Userid.Visible = False
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = lv_Userid.ListItems.Add(, , rsFilter!userid)
        mitem.SubItems(1) = rsFilter!posisi
        'mitem.SubItems(2) = rsFilter!posisi
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
End Sub

Public Sub reload_Member()
    lv_memberId.ListItems.Clear
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("select * from v_member where memberid like '%" & txt_memberID.Text & "%'")
    
    If rsFilter.EOF Then
        lv_memberId.Visible = False
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = lv_memberId.ListItems.Add(, , rsFilter!memberid)
        mitem.SubItems(1) = rsFilter!nama
        'mitem.SubItems(2) = rsFilter!posisi
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
End Sub

Function check_memberID(temp_memberID As String) As Boolean
    Dim rslogin As ADODB.Recordset
    Set rslogin = con.Execute("select * from v_member where memberid = '" & txt_memberID.Text & "'")
    
    If rslogin.EOF Then check_memberID = False Else check_memberID = True
End Function

Function check_userID(temp_userID As String) As Boolean
    
    Dim rslogin As ADODB.Recordset
    Set rslogin = con.Execute("select * from tblogin where userid = '" & temp_userID & "'")

    If rslogin.EOF Then check_userID = False Else check_userID = True

End Function

Private Sub openReport(file_name As String, db_column As String, report_type As String, flag_id As Boolean)
    'cr.connect = "Provider=MSDASQL.1;Pwd=" & Setting_Object("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object("DB_Id") & ";Data Source=Data"
    cr.connect = "Provider=MSDASQL.1;Pwd=yuyu;Persist Security Info=True;User ID=root;Data Source=parkir"
    
    cr.ReportFileName = App.Path + "\" + file_name
    
    If report_type = "ONE_DAY" Then
        cr.SelectionFormula = "{" & db_column & "}= #" & Format(dt_start.Value, "yyyy-MM-dd") & "#"
        
        cr.Formulas(0) = "tgl1='" & "Tanggal : " & Format(dt_start.Value, "dd/MM/yyyy") & "'"
        cr.Formulas(1) = "petugas ='" & "Pengawas : " & txt_Userid.Text & "'"
        
    ElseIf report_type = "DURATION" Then
        cr.SelectionFormula = "{" & db_column & "}>= #" & Format(dt_start.Value, "yyyy-MM-dd") & "# and {" & db_column & "}<= #" & Format(dt_end.Value, "yyyy-MM-dd") & "#"
         
        cr.Formulas(0) = "tgl1='" & "Dari : " & Format(dt_start.Value, "dd/MM/yyyy") & "'"
        cr.Formulas(1) = "tgl2='" & "Sampai : " & Format(dt_end.Value, "dd/MM/yyyy") & "'"
        cr.Formulas(2) = "petugas='" & "Pengawas : " & txt_Userid.Text & "'"
    End If
    
    If flag_id = True Then
        If file_name = "laporanharian3.rpt" Or file_name = "laporanharian4.rpt" Then
            cr.SelectionFormula = cr.SelectionFormula & "and {laporan.userid} = '" & txt_Userid.Text & "'"
        ElseIf file_name = "laporanmember.rpt" Then
            cr.SelectionFormula = cr.SelectionFormula & "and {tbtransaksi.memberid} = '" & txt_memberID.Text & "'"
        End If
    End If
    
    cr.WindowState = crptMaximized
    cr.RetrieveDataFiles
    cr.Action = 1
    cr.reset
End Sub
