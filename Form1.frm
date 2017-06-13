VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_Main 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Layar Utama"
   ClientHeight    =   10950
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btn_PengaturanMember 
      Caption         =   "Pengaturan Member"
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
      Left            =   17880
      TabIndex        =   25
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton btn_PengaturanUser 
      Caption         =   "Pengaturan User"
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
      Left            =   16080
      TabIndex        =   24
      Top             =   7440
      Width           =   1575
   End
   Begin VB.PictureBox pic_Foto 
      Height          =   5182
      Left            =   6600
      ScaleHeight     =   5115
      ScaleWidth      =   6840
      TabIndex        =   14
      Top             =   5760
      Width           =   6900
      Begin VB.Label lbl_close 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6600
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton btn_CameraSettings 
      Caption         =   "Pengaturan Kamera"
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
      Left            =   14280
      TabIndex        =   22
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Timer scanner_Timeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   13680
      Top             =   2160
   End
   Begin VB.CommandButton btn_Logout 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   20
      Top             =   9600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   13680
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btn_AmbilFoto 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ambil Foto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   960
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton btn_Exit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tutup Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3600
      Picture         =   "Form1.frx":0F5A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6240
      Top             =   360
   End
   Begin VB.CommandButton cmd_BukaDarurat 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buka Portal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3600
      Picture         =   "Form1.frx":19FE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   13680
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txt_kode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      TabIndex        =   0
      Top             =   480
      Width           =   5295
   End
   Begin VB.Frame frame_1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   6015
      Begin VB.CommandButton btn_Laporan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Laporan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   720
         Picture         =   "Form1.frx":2D21
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton btn_BukaFoto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buka Foto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3360
         Picture         =   "Form1.frx":3E31
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton btn_BukaNormal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Karcis Mobil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form1.frx":4699
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   9120
         Width           =   6015
      End
      Begin VB.Label lbl_Jam 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         TabIndex        =   9
         Top             =   9600
         Width           =   2175
      End
      Begin VB.Label lbl_Tgl 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "14-02-2016"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   9600
         Width           =   2775
      End
   End
   Begin VB.Frame Frame_Scan 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Data Mobil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   14280
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton btn_TutupFrameMobil 
         Caption         =   "Tutup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   12
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label lbl_Status 
         BackStyle       =   0  'Transparent
         Caption         =   "Status : Masih Berlaku"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   240
         TabIndex        =   18
         Top             =   3720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lbl_Perusahaan 
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan : abcdefghijklmnopqrstuvwxyz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lbl_Nama 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama : abcdefghijklmnopqrstuvwxyz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lbl_MemberID 
         BackStyle       =   0  'Transparent
         Caption         =   "Member ID : 12345678901234567890"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lbl_Warning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Data Tidak Ditemukan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   3975
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   4695
      End
   End
   Begin VB.Image img_empty 
      Height          =   2655
      Left            =   6600
      Picture         =   "Form1.frx":5735
      Top             =   5760
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.Image imgPlaceHolder 
      Height          =   5182
      Left            =   6600
      Top             =   360
      Width           =   6900
   End
   Begin VB.Label lbl_Uang 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Jumlah Uang :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14280
      TabIndex        =   21
      Top             =   8760
      Width           =   5295
   End
   Begin VB.Label lbl_Petugas 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Petugas : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14280
      TabIndex        =   19
      Top             =   8160
      Width           =   5295
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Requires a reference to:
'
'   ActiveMovie control type library (quartz.dll).
'

Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const MASKBORDERLESS = Not (WS_BORDER Or WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)
Private Const MASKBORDERMIN = Not (WS_DLGFRAME Or WS_SYSMENU Or WS_THICKFRAME)

'FILTER_STATE values, should have been defined in Quartz.dll,
'but another item Microsoft left out.
Private Enum FILTER_STATE
    State_Stopped = 0
    State_Paused = 1
    State_Running = 2
End Enum

Private Const E_FAIL As Long = &H80004005

'These are "scripts" followed by BuildGraph() below to create a
'DirectShow FilterGraph for webcam viewing.
'
'FILTERLIST is incomplete, and must be prepended with the name
'of your webcam's Video Capture Source filter.  Since there may
'be multiples, FILTERLIST begins with "~Capture" which is used
'when BuildGraph() interprets this script to select one having
'a pin named "Capture".
Private Const FILTERLIST As String = _
        "~Capture|" _
      & "AVI Decompressor|" _
      & "Color Space Converter|" _
      & "Video Renderer"
Private Const CONNECTIONLIST As String = _
        "Capture~XForm In|" _
      & "XForm Out~Input|" _
      & "XForm Out~VMR Input0"

Private fgmVidCap As QuartzTypeLib.FilgraphManager 'Not "Is Nothing" means camera is previewing.
Private bv2VidCap As QuartzTypeLib.IBasicVideo2
Private vwVidCap As QuartzTypeLib.IVideoWindow
Private SelectedCamera As Integer '-1 means none selected.
Private InsideWidth As Double
Private AspectRatio As Double

Public kode_Transaksi As String
Public last_Member, temp_Member As String
Public total_uang As Long
Dim camera_ready As Boolean

Sub take_pic(folder_location As String, namafoto As String)
    On Error GoTo error_handler2
    If camera_ready = True Then
        Const PauseWaitMs As Long = 16
        Const biSize = 40 'BITMAPINFOHEADER and not BITMAPV4HEADER, etc. but we don't get those.
        Dim State As FILTER_STATE
        Dim Size As Long
        Dim DIB() As Long
        Dim hBitmap As Long
        Dim Pic As StdPicture
        
        With fgmVidCap
            .Pause
            Do
                .GetState PauseWaitMs, State
            Loop Until State = State_Paused Or Err.Number = E_FAIL
            If Err.Number = E_FAIL Then
                MsgBox "Failed to pause webcam preview for snapshot!", _
                       vbOKOnly Or vbExclamation
                Exit Sub
            End If
        
            With bv2VidCap
                'Estimate size.  Correct for 32-bit RGB and generous
                'for anything with fewer bits per pixel, compressed,
                'or palette-ized (we hope).
                Size = biSize + .VideoWidth * .VideoHeight
                ReDim DIB(Size - 1)
                Size = Size * 4 'To bytes.
                .GetCurrentImage Size, DIB(0)
            End With
            
            .Run
        End With
        
        hBitmap = LongDIB2HBitmap(DIB)
        If hBitmap <> 0 Then
            Set Pic = HBitmap2Picture(hBitmap, 0)
            If Not Pic Is Nothing Then
                With pic_Foto
                    .AutoRedraw = True
                    .PaintPicture Pic, 0, 0, .ScaleWidth, .ScaleHeight
                     'Call SavePicture(.Image, App.Path & "\Foto\" & Left(kode_Transaksi, 1) & Format(Now, "yyyy-MM-dd-h-mm-ss") & ".jpg")
                    Call SavePicture(.Image, App.Path & folder_location & namafoto)
                    .AutoRedraw = False
                End With
            End If
            DeleteObject hBitmap
        End If
        lbl_close.Visible = True
        FindFiles (folder_location)
    End If
    Exit Sub
error_handler2:
    Debug.Print "Gambar tidak bisa diambil"
    camera_ready = False
End Sub

Private Sub btn_AmbilFoto_Click()
    Timer1.Enabled = False
    DoEvents
    Call take_pic("\Foto\", Left(kode_Transaksi, 1) & Format(Now, "yyyy-MM-dd-h-mm-ss") & ".jpg")
    Timer1.Enabled = True
End Sub

Private Sub btn_BukaFoto_Click()
    Timer1.Enabled = False
    DoEvents
    OpenShowFile
    Timer1.Enabled = True
End Sub

Private Sub btn_BukaNormal_Click()
    Timer1.Enabled = False
    DoEvents
    If (MsgBox("Buka gerbang?", vbYesNo, "Open") = vbYes) Then
        con.Execute ("insert into tbtransaksi values ('" & kode_Transaksi & "','" & Format(Now, "yyyy-MM-dd") & "','" & Format(Now, "HH:mm:ss") & "','" & "3000" & "','" & "Karcis" & "','','" & username & "')")
        Buka_Pintu
        nextFaktur
        total_uang = total_uang + 3000
        lbl_Uang.Caption = "Jumlah Uang : " & Format(total_uang, "###,###,##0")
    End If
    Timer1.Enabled = True
End Sub

Private Sub btn_CameraSettings_Click()
    If isMaster = True Then
    
        Timer1.Enabled = False
        DoEvents
        Dim StartResult As Integer
        
        Form_ListCameras.Show vbModal, Me
        If Form_ListCameras.Oked Then
            
            StopCamera
            StartResult = StartCamera(Form_ListCameras.CameraName)
            If StartResult < 0 Then
                
                SaveSettings (Form_ListCameras.CameraName)
                
            Else
                MsgBox "This doesn't seems to be a valid webcam:" & vbNewLine _
                     & vbNewLine _
                     & Form_ListCameras.CameraName & vbNewLine _
                     & vbNewLine _
                     & "BuildGraph error " & CStr(Error), _
                       vbOKOnly Or vbInformation
                'Try to go back to previous camera.
                
            End If
            
        End If
        picStrech
    Else
        LoadSettings
    End If
End Sub

Private Sub btn_Exit_Click()
    Timer1.Enabled = False
    DoEvents
    If isMaster = False Then
        If (MsgBox("Cetak Laporan?", vbYesNo, "Cetak Laporan") = vbYes) Then
            btn_Laporan_Click
            Timer1.Enabled = False
            DoEvents
        End If
    End If
    
    If (MsgBox("Yakin Keluar?", vbYesNo, "Keluar") = vbYes) Then
        Dim Form As Form
        For Each Form In Forms
        Unload Form
        Set Form = Nothing
        Next Form
    Else
        Timer1.Enabled = True
    End If
    
End Sub

Private Sub btn_Laporan_Click()
    Timer1.Enabled = False
    DoEvents
    If isMaster = False Then
    
        Dim jumlah_karcis, jumlah_member, jumlah_buka, jumlah_darurat As Integer
        Dim jumlah_uang As Long
        Dim rstrans As ADODB.Recordset
        
        Set rstrans = con.Execute("select count(kode) as sum1, sum(bayar) as sum2 from tbtransaksi where tanggal = '" & Format(Now, "yyyy-MM-dd") & "' and status = 'Karcis' and userid = '" & username & "'")
        jumlah_karcis = rstrans!sum1
        'jumlah_uang = rstrans!sum2
        If rstrans!sum2 > 0 Then
            jumlah_uang = rstrans!sum2
        Else
            jumlah_uang = 0
        End If
        Set rstrans = con.Execute("select count(kode) as sum3 from tbtransaksi where tanggal = '" & Format(Now, "yyyy-MM-dd") & "' and status = 'Member' and userid = '" & username & "'")
        jumlah_member = rstrans!sum3
        
        Set rstrans = con.Execute("select count(kode) as sum4 from tbbuka where tanggal = '" & Format(Now, "yyyy-MM-dd") & "' and userid = '" & username & "'")
        jumlah_darurat = rstrans!sum4
        jumlah_buka = jumlah_darurat + jumlah_karcis + jumlah_member
        
        
        
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.Font = "dotumche"
        Printer.FontSize = 17
        Printer.FontBold = True
        Printer.Print Tab(8); "REKAP PARKIR";
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.Print Tab(4); "                                    "
        Printer.Print Tab(4); "Nama    : "; username
        Printer.Print Tab(4); "Tanggal : "; Format(Now, "dd/MM/yyyy")
        Printer.Print Tab(4); "Jam     : "; Format(Now, "HH:mm:ss")
        Printer.Print Tab(4); "                                    ";
        Printer.Print Tab(4); "Jumlah Karcis         : "; jumlah_karcis; "lembar"
        Printer.Print Tab(4); "Jumlah Uang           : Rp."; Format(jumlah_uang, "###,###,##0")
        If jumlah_member > 0 Then Printer.Print Tab(4); "Jumlah Member lewat   : "; jumlah_member; "kali"
        If jumlah_darurat > 0 Then Printer.Print Tab(4); "Jumlah dibuka darurat : "; jumlah_darurat; "kali"
        Printer.Print Tab(4); "Jumlah Pintu Dibuka   : "; jumlah_buka; "kali"
        
        Printer.EndDoc
    Else

        Form_Laporan.Show (1)
    End If
    DoEvents
    If Timer1.Enabled = False Then
        Timer1.Enabled = True
    End If
End Sub

Private Sub btn_Logout_Click()
    username = ""
    Timer1.Enabled = False
    DoEvents
    Unload Me
    con.Close
    con2.Close
    Form_Login.Show (1)
End Sub

Private Sub btn_PengaturanMember_Click()
    Timer1.Enabled = False
    DoEvents
    form_Member.Show vbModal, Me
End Sub

Private Sub btn_PengaturanUser_Click()
    Timer1.Enabled = False
    DoEvents
    Form_User.Show vbModal, Me
End Sub

Private Sub btn_TutupFrameMobil_Click()
    Frame_Scan.Visible = False
End Sub

Sub Buka_Pintu()
    MSComm1.Output = "X"
    Sleep 1500
    MSComm1.Output = "P"
    'txt_kode.SetFocus
End Sub

Private Sub cmd_BukaDarurat_Click()
    Timer1.Enabled = False
    DoEvents
    Form_Gate.Show (1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'txt_kode.SetFocus
    'txt_kode.Text = KeyAscii
    'txt_kode.SelectionStart = txtSearch.Text.Length
    'e.Handled = True
End Sub

Private Sub Form_Load()

    MSComm1.CommPort = 1
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.InputLen = 0
    MSComm1.RThreshold = 1
    MSComm1.PortOpen = True
    
    'VideoGrabberVB61.VideoDevice = 0
    'VideoGrabberVB61.StartPreview
    
    lbl_Tgl.Caption = Format(Now, "dd-MM-yyyy")
    lbl_Jam.Caption = Format(Now, "HH:mm:ss")
    txt_kode.Text = ""
    lbl_Petugas.Caption = "Petugas : " & username
    
    
    
    Dim namafile, file_data, huruf As String
    Dim angka As Long
    namafile = App.Path & "\faktur.txt"
    Open namafile For Input As #1
    While Not EOF(1)
        Input #1, file_data
        'file_data = data
        huruf = Left(file_data, 1)
        angka = Val(Mid(file_data, 2, 20))
        kode_Transaksi = huruf + CStr(angka + 1)
    Wend
    Close #1
    
    Dim StartResult As Integer
    
    InsideWidth = pic_Foto.Width - ScaleX(2, vbPixels, ScaleMode)
    LoadSettings
    pic_Foto.Picture = img_empty.Picture
    picStrech
    
    getTotalUang
    lbl_Uang.Caption = "Jumlah Uang : " & Format(total_uang, "###,###,##0")
    
    If isMaster Then
        btn_CameraSettings.Visible = True
    Else
        'btn_CameraSettings.Visible = False
        btn_CameraSettings.Caption = "Refresh Camera"
    End If
    If isMaster Then btn_PengaturanUser.Visible = True Else btn_PengaturanUser.Visible = False
    If isMaster Then btn_PengaturanMember.Visible = True Else btn_PengaturanMember.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    StopCamera
    Set Form_Main = Nothing
End Sub

Sub tutup_foto()
    lbl_close.Visible = False
    pic_Foto.Picture = img_empty.Picture
    picStrech
End Sub

Private Sub lbl_close_Click()
    tutup_foto
End Sub

Private Sub scanner_Timeout_Timer()
    last_Member = ""
    scanner_Timeout.Enabled = False
End Sub

Private Sub Timer1_Timer()
    lbl_Tgl.Caption = Format(Now, "dd-MM-yyyy")
    lbl_Jam.Caption = Format(Now, "HH:mm:ss")
    txt_kode.Text = ""
    txt_kode.SetFocus
    
End Sub

Private Sub txt_kode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'control+v paste
    If Shift = vbCtrlMask And (Chr(KeyCode) = "v" Or Chr(KeyCode) = "V") Then
        txt_kode.Locked = True
    Else
        txt_kode.Locked = False
    End If
    
    If Len(txt_kode.Text) < 1 Then
        Timer1.Enabled = False
        DoEvents
        Timer1.Enabled = True
    End If
    
    If KeyCode = 13 And txt_kode.Text <> "" Then
        temp_Member = txt_kode.Text
        'temp_Member = Replace(temp_Member, Chr(39), "")
        txt_kode.Text = ""
        If temp_Member = last_Member Then
            Frame_Scan.Visible = True
            lbl_MemberID.Visible = True
            lbl_Nama.Visible = True
            lbl_Perusahaan.Visible = False
            lbl_Status.Visible = False
            lbl_Warning.Visible = False
            
            lbl_MemberID.Caption = "Member ID : " & last_Member
            lbl_Nama.Caption = "Silahkan Lewat"
            
            DoEvents
                
            Buka_Pintu
        
        Else
            If Len(temp_Member) = 8 Then
                scanBarcode (True)
            Else
                scanBarcode (False)
            End If
            
        End If
        'MsgBox txt_kode.Text
    End If
End Sub

Private Sub txt_kode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 65 To 90, 48 To 57, 8 ' A-Z, 0-9 and backspace
        'Let these key codes pass through
        Case 97 To 122, 8 'a-z and backspace
        'Let these key codes pass through
        Case Else
        'All others get trapped
        KeyAscii = 0 ' set ascii 0 to trap others input
    End Select
End Sub

Private Sub txt_kode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'right click paste
   If Button = vbRightButton Then
        txt_kode.Locked = True
    Else
        txt_kode.Locked = False
    End If
End Sub

Private Sub OpenShowFile()
    On Error GoTo OpenShowFileError
    Dim FName As String, FNumb As Integer
    Dim FileContents As String
    cd.CancelError = True
    cd.Filter = "All Images (*.jpg, *.bmp, *.gif)|*.jpg;*.bmp;*.gif|All Files (*.*)|*.*"
    cd.InitDir = App.Path & "\Foto\"
    cd.filename = vbNullString
    cd.ShowOpen
    If cd.filename = vbNullString Then Exit Sub
    
    pic_Foto.Picture = LoadPicture(cd.filename)
    lbl_close.Visible = True
    picStrech
    
    Exit Sub
OpenShowFileError:
    If Err.Number = 32755 Then Exit Sub 'user pressed cancel
End Sub

Public Sub nextFaktur()
    Dim namafile, huruf As String
    Dim angka As Long
    Me.Enabled = True
    huruf = Left(kode_Transaksi, 1)
    angka = Val(Mid(kode_Transaksi, 2, 20))
    
    namafile = App.Path & "\faktur.txt"
    Open namafile For Output As #1
    Print #1, kode_Transaksi
    Close #1
    
    kode_Transaksi = huruf + CStr(angka + 1)
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

Sub FindFiles(folder_Name)
    Dim strFileName As String
    Dim strOldestFile As String
    Dim strFolder As String
    Dim iFileCount As Integer
    Dim i As Integer
    Dim dt As Date
    
'    strFolder = App.Path & "\Foto"
'    strFileName = Dir$(strFolder & "\*")
'    strOldestFile = strFileName
'    dt = FileDateTime(strFolder & "\" & strFileName)
'    Do Until strFileName = ""
'    Debug.Print strFileName, FileDateTime(strFolder & "\" & strFileName),
'        If (FileDateTime(strFolder & "\" & strFileName)) < dt Then
'            dt = FileDateTime(strFolder & "\" & strFileName)
'            strOldestFile = strFileName
'        End If
'        i = i + 1
'        strFileName = Dir$()
'    Loop
'    If i > 100 Then
'        Kill strFolder & "\" & strOldestFile
'    End If
    strFolder = App.Path & folder_Name
    strFileName = Dir$(strFolder & "*")
    strOldestFile = strFileName
    dt = FileDateTime(strFolder & strFileName)
    Do Until strFileName = ""
    Debug.Print strFileName, FileDateTime(strFolder & strFileName),
        If (FileDateTime(strFolder & strFileName)) < dt Then
            dt = FileDateTime(strFolder & strFileName)
            strOldestFile = strFileName
        End If
        i = i + 1
        strFileName = Dir$()
    Loop
    If i > 100 Then
        Kill strFolder & strOldestFile
    End If
End Sub

Sub scanBarcode(DB_toggle As Boolean)
    Dim rsMember As ADODB.Recordset
    
    If DB_toggle = True Then
        Set rsMember = con2.Execute("Select * from member where member_id = '" & temp_Member & "'")
    Else
        Set rsMember = con.Execute("Select * from tbmember where memberid = '" & temp_Member & "'")
    End If
    
    If Not rsMember.EOF Then
        lbl_MemberID.Caption = "Member ID : " & temp_Member
        lbl_Nama.Caption = "Nama : " & rsMember!nama
        lbl_Perusahaan.Caption = "Perusahaan : " & rsMember!perusahaan
        
        Dim b As String
        
        If rsMember!status = 1 Then b = "Masih Berlaku" Else b = "Tidak Berlaku"
        
        lbl_Status = "Status : " & b
        
        Frame_Scan.Visible = True
        lbl_MemberID.Visible = True
        lbl_Nama.Visible = True
        lbl_Perusahaan.Visible = True
        lbl_Status.Visible = True
        lbl_Warning.Visible = False
        
        If b = "Tidak Berlaku" Then
            'red
            lbl_Status.ForeColor = &HFF&
        Else
            'black
            lbl_Status.ForeColor = &H0&
            last_Member = temp_Member
            scanner_Timeout.Enabled = True
            
            con.Execute ("insert into tbtransaksi values ('" & kode_Transaksi & "','" & Format(Now, "yyyy-MM-dd") & "','" & Format(Now, "HH:mm:ss") & "','" & "0" & "','" & "Member" & "','" & temp_Member & "','" & username & "')")
            
            DoEvents
            
            Buka_Pintu
            nextFaktur
            
        End If
        
        

    Else
        Frame_Scan.Visible = True
        lbl_MemberID.Caption = "Member ID : " & temp_Member
        lbl_MemberID.Visible = True
        lbl_Warning.Visible = True
        lbl_Nama.Visible = False
        lbl_Perusahaan.Visible = False
        lbl_Status.Visible = False
        
    End If
    
    temp_Member = ""
    
End Sub

Private Sub getTotalUang()
    Dim rssum As ADODB.Recordset
    Set rssum = con.Execute("select sum(bayar) as sum1 from tbtransaksi where userid = '" & username & "' and tanggal = '" & Format(Now, "yyyy-MM-dd") & "'")
    
    If rssum!sum1 > 0 Then
        total_uang = rssum!sum1
    Else
        total_uang = 0
    End If
    
End Sub

Private Function BuildGraph( _
    ByVal FGM As QuartzTypeLib.FilgraphManager, _
    ByVal Filters As String, _
    ByVal Connections As String) As Integer
    'Returns -1 on success, or FilterIndex when not found, or
    'ConnIndex + 100 when a pin of the connection not found.
    '
    'Filters:
    '
    '   A string with Filter Name values separated by "|" delimiters
    '   and optionally each of these can be followed by one required
    '   Pin Name value separated by a "~" delimiter for use as a tie
    '   breaker when there might be multiple filters with the same
    '   Name value.
    '
    'Connections:
    '
    '   A string with a list of output pins to be connected to
    '   input pins.  Each pin-pair is separated by "|" delimiters
    '   and each pair has out and in pins separated by a "~"
    '   delimiter.  The pin-pairs should be one less than the number
    '   of filters.
    On Error GoTo Error_Handler
    Dim FilterNames() As String
    Dim FilterIndex As Integer
    Dim FilterParts() As String
    Dim FoundFilter As Boolean
    Dim rfiEach As QuartzTypeLib.IRegFilterInfo
    Dim fiFilters() As QuartzTypeLib.IFilterInfo
    Dim Conns() As String
    Dim ConnIndex As Integer
    Dim ConnParts() As String
    Dim piEach As QuartzTypeLib.IPinInfo
    Dim piOut As QuartzTypeLib.IPinInfo
    Dim piIn As QuartzTypeLib.IPinInfo
    
    'Setup for filter script processing.
    FilterNames = Split(UCase$(Filters), "|")
    ReDim fiFilters(UBound(FilterNames))
    
    'Find and add filters.
    For FilterIndex = 0 To UBound(FilterNames)
        FilterParts = Split(FilterNames(FilterIndex), "~")
        For Each rfiEach In FGM.RegFilterCollection
            If UCase$(rfiEach.Name) = FilterParts(0) Then
                rfiEach.Filter fiFilters(FilterIndex)
                If UBound(FilterParts) > 0 Then
                    For Each piEach In fiFilters(FilterIndex).Pins
                        If UCase$(piEach.Name) = FilterParts(1) Then
                            FoundFilter = True
                            Exit For
                        End If
                    Next
                Else
                    FoundFilter = True
                    Exit For
                End If
            End If
        Next
        If FoundFilter Then
            FoundFilter = False
        Else
            BuildGraph = FilterIndex
            Exit Function 'Error result will be 0, 1, etc.
        End If
    Next
    BuildGraph = -1
    
    'Setup for connection script processing.
    Conns = Split(UCase$(Connections), "|")
    FilterIndex = 0
    
    'Find and connect pins.
    For ConnIndex = 0 To UBound(Conns)
        ConnParts = Split(Conns(ConnIndex), "~")
        For Each piEach In fiFilters(FilterIndex).Pins
            If UCase$(piEach.Name) = ConnParts(0) Then
                Set piOut = piEach
                Exit For
            End If
        Next
        For Each piEach In fiFilters(FilterIndex + 1).Pins
            If UCase$(piEach.Name) = ConnParts(1) Then
                Set piIn = piEach
                Exit For
            End If
        Next
        If piOut Is Nothing Or piIn Is Nothing Then
            'Error, missing a pin.
            BuildGraph = ConnIndex + 100 'Error result will be 100, 101, etc.
            Exit Function
        End If
        piOut.ConnectDirect piIn
        FilterIndex = FilterIndex + 1
    Next
    
    Exit Function
    
Error_Handler:
    BuildGraph = 3
End Function

Private Sub DeselectFailedCamera(ByVal Error As Long)
    Dim CameraName As String
    
    SelectedCamera = -1
    MsgBox "Selected camera failed, may not be connected:" & vbNewLine _
         & vbNewLine _
         & CameraName & vbNewLine _
         & vbNewLine _
         & "BuildGraph error " & CStr(Error), _
           vbOKOnly Or vbInformation
End Sub



Private Sub LoadSettings()
    Dim F As Integer
    Dim C As Integer
    Dim CameraName As String

    SelectedCamera = -1 'None.
    On Error Resume Next
    GetAttr "CameraSettings.txt"
    If Err.Number = 0 Then
        On Error GoTo 0
        F = FreeFile(0)
        Open "CameraSettings.txt" For Input As #F
        Input #F, SelectedCamera
        Do Until EOF(F)
            Input #F, CameraName
            StartCamera (CameraName)
            C = C + 1
        Loop
        Close #F
    End If
End Sub

Private Sub SaveSettings(temp_camName As String)
    Dim F As Integer
    Dim C As Integer

    F = FreeFile(0)
    Open "CameraSettings.txt" For Output As #F
    Write #F, 0
    Write #F, temp_camName
    Close #F
End Sub

Private Function StartCamera(ByVal CamName As String) As Integer
    'Returns -1 on success, or BuildGraph() error on failures.
    
    Set fgmVidCap = New QuartzTypeLib.FilgraphManager
    'Tack camera name onto FILTERLIST and try to start it.
    StartCamera = BuildGraph(fgmVidCap, CamName & FILTERLIST, CONNECTIONLIST)
    If StartCamera >= 0 Then
        camera_ready = False
        Exit Function
    Else
        camera_ready = True
    End If
    
    Set bv2VidCap = fgmVidCap
    With bv2VidCap
        AspectRatio = CDbl(.VideoHeight) / CDbl(.VideoWidth)
    End With
    
    Set vwVidCap = fgmVidCap
    With vwVidCap
        .FullScreenMode = False
        .Left = ScaleX(imgPlaceHolder.Left, ScaleMode, vbPixels)
        .Top = ScaleY(imgPlaceHolder.Top, ScaleMode, vbPixels)
        .Width = ScaleX(InsideWidth, ScaleMode, vbPixels) + 2
        .Height = ScaleY(InsideWidth * AspectRatio, ScaleMode, vbPixels) + 2
        pic_Foto.Height = InsideWidth * AspectRatio + ScaleY(2, vbPixels, ScaleMode)
        imgPlaceHolder.Visible = False
        .WindowStyle = .WindowStyle And MASKBORDERMIN
        .Owner = hWnd
        .Visible = True
    End With
    
    StartCamera = -1
    fgmVidCap.Run
    'picStrech
End Function

Private Sub StopCamera()
    Const StopWaitMs As Long = 40
    Dim State As FILTER_STATE
    
    If Not fgmVidCap Is Nothing Then
        With fgmVidCap
            .Stop
            Do
                .GetState StopWaitMs, State
            Loop Until State = State_Stopped Or Err.Number = E_FAIL
        End With
        If Not vwVidCap Is Nothing Then
            With vwVidCap
                .Visible = False
                .Owner = 0
            End With
            Set vwVidCap = Nothing
        End If
        Set bv2VidCap = Nothing
        Set fgmVidCap = Nothing
    End If
    imgPlaceHolder.Visible = True
    'cmdSnap.Enabled = False
End Sub


