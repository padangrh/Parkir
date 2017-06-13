VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Review 
   Caption         =   "Review"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   19230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_CHDir 
      Caption         =   "Change Folder Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13080
      TabIndex        =   24
      Top             =   6120
      Width           =   4815
   End
   Begin VB.TextBox txt_FotoID 
      Enabled         =   0   'False
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
      Left            =   6360
      TabIndex        =   22
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txt_Keterangan 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   5520
      Width           =   5535
   End
   Begin VB.TextBox txt_UserID 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   18
      Top             =   7320
      Width           =   2415
   End
   Begin VB.TextBox txt_Jam 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   16
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox txt_Tanggal 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   14
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox txt_Kode 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   12
      Top             =   5520
      Width           =   2415
   End
   Begin VB.PictureBox pic_Foto 
      Height          =   5055
      Left            =   12120
      ScaleHeight     =   4995
      ScaleWidth      =   6840
      TabIndex        =   11
      Top             =   720
      Width           =   6900
   End
   Begin VB.CheckBox chk_Filter 
      Caption         =   "Filter"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "By"
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   11535
      Begin VB.TextBox txt_Search 
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
         Left            =   8880
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox chk_Sampai 
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin MSComCtl2.DTPicker dt_start 
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   360
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   98238467
         CurrentDate     =   42810
      End
      Begin MSComCtl2.DTPicker dt_end 
         Height          =   495
         Left            =   5760
         TabIndex        =   5
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   98238467
         CurrentDate     =   42810
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User"
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
         Left            =   8040
         TabIndex        =   8
         Top             =   360
         Width           =   1095
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
         Left            =   4680
         TabIndex        =   7
         Top             =   360
         Width           =   1455
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
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView LV_Buka 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tanggal"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Jam"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "User Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Keterangan"
         Object.Width           =   7408
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Foto ID"
         Object.Width           =   4057
      EndProperty
   End
   Begin VB.Image img_empty 
      Height          =   2655
      Left            =   12120
      Picture         =   "Form_Review.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Foto ID"
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
      Left            =   4920
      TabIndex        =   23
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   4800
      TabIndex        =   21
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Jam"
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
      Left            =   360
      TabIndex        =   17
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label6 
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
      Left            =   360
      TabIndex        =   15
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode"
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
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Review pintu dibuka darurat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form_Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic_path As String

Private Sub btn_CHDir_Click()
    Dim temp_path As String
    temp_path = BrowseForFolder(hWnd, "Please select the picture folder.")
    If temp_path <> "" Then pic_path = temp_path
End Sub

Private Sub chk_Filter_Click()
    reload_List
End Sub

Private Sub chk_Sampai_Click()
    If chk_Sampai = 1 Then
        dt_end.Enabled = True
    Else
        dt_end.Enabled = False
    End If
    reload_List
End Sub

Private Sub dt_end_Change()
    reload_List
End Sub

Private Sub dt_start_Change()
    reload_List
End Sub

Private Sub Form_Load()
    dt_start.Value = Now
    dt_end.Value = Now
    pic_path = App.Path & "\Simpan"
    reload_List
    pic_Foto.Picture = img_empty.Picture
    picStrech
End Sub


Public Sub reload_List()
'pindahan generate list barang
    LV_Buka.ListItems.Clear
    'list_nama.Visible = True
    Dim rsFilter As ADODB.Recordset
    Dim StringQuery As String
    
    StringQuery = "Select * from tbbuka"
    
    If chk_Filter.Value = 1 Then
        If chk_Sampai.Value = 1 Then
            StringQuery = "Select * from tbbuka where tanggal >= '" & Format(dt_start, "yyyy-mm-dd") & "' and tanggal <= '" & Format(dt_end, "yyyy-mm-dd") & "'"
        Else
            StringQuery = "Select * from tbbuka where tanggal = '" & Format(dt_start, "yyyy-mm-dd") & "'"
        End If
        StringQuery = StringQuery & " and userid like '%" & txt_Search.Text & "%'"
    End If
    
    Set rsFilter = con.Execute(StringQuery)
    
    If rsFilter.EOF Then
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = LV_Buka.ListItems.Add(, , rsFilter!kode)
        mitem.SubItems(1) = rsFilter!tanggal
        mitem.SubItems(2) = rsFilter!jam
        mitem.SubItems(3) = rsFilter!userid
        mitem.SubItems(4) = rsFilter!keterangan
        mitem.SubItems(5) = rsFilter!fotoid
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
'end pindahan list barang
End Sub

Private Sub LV_Buka_DblClick()
    Dim fso As FileSystemObject
    
    txt_Kode.Text = LV_Buka.SelectedItem.Text
    txt_Tanggal.Text = LV_Buka.SelectedItem.SubItems(1)
    txt_Jam.Text = LV_Buka.SelectedItem.SubItems(2)
    txt_Userid = LV_Buka.SelectedItem.SubItems(3)
    txt_Keterangan = LV_Buka.SelectedItem.SubItems(4)
    txt_FotoID = LV_Buka.SelectedItem.SubItems(5)
    
    If (Right(pic_path, 1) <> "\") Then pic_path = pic_path & "\"
    
    Set fso = New FileSystemObject
    If fso.FileExists(pic_path & txt_FotoID) Then
        'load picture
        pic_Foto.Picture = LoadPicture(pic_path & txt_FotoID)
    Else
        pic_Foto.Picture = img_empty.Picture
        picStrech
    End If
    
End Sub

Private Sub txt_Search_Change()
    reload_List
End Sub

Sub picStrech()
    pic_Foto.ScaleMode = 3
    pic_Foto.AutoRedraw = True
    pic_Foto.PaintPicture pic_Foto.Picture, _
        0, 0, pic_Foto.ScaleWidth, pic_Foto.ScaleHeight, _
        0, 0, _
        pic_Foto.Picture.Width / 26.46, _
        pic_Foto.Picture.Height / 26.46
    pic_Foto.Picture = pic_Foto.Image
End Sub

