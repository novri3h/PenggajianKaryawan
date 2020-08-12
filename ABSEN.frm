VERSION 5.00
Begin VB.Form ABSEN 
   Caption         =   "FORM ABSEN"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   3330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox NAMA 
      Height          =   400
      Left            =   1680
      TabIndex        =   10
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1440
      Top             =   840
   End
   Begin VB.CommandButton JAMKELUAR 
      Caption         =   "KELUAR"
      Height          =   400
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3060
   End
   Begin VB.CommandButton JAMMASUK 
      Caption         =   "MASUK"
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   3060
   End
   Begin VB.TextBox NIP 
      Height          =   400
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA "
      Height          =   405
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JAM NORMAL"
      Height          =   405
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label NORMAL 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "08:00:00"
      Height          =   405
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label WAKTU 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WAKTU"
      Height          =   405
      Left            =   1680
      TabIndex        =   7
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WAKTU"
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIP"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label TANGGAL 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TANGGAL"
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1500
   End
End
Attribute VB_Name = "ABSEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Label2 = Format(CDate(WAKTU) - CDate(NORMAL), "HH:MM:SS")
End Sub

Private Sub Form_Load()
TANGGAL = Date
End Sub


Private Sub JAMKELUAR_Click()
If NIP = "" Or NAMA = "" Then
    MsgBox "NIP DAN NAMA HARUS DIISI"
    If NIP = "" Then
        NIP.SetFocus
        Exit Sub
    ElseIf NAMA = "" Then
        NAMA.SetFocus
        Exit Sub
    End If
End If

Call BukaDB
RSABSEN.Open "SELECT * FROM ABSEN WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "' AND MASUK<>NULL", Conn
If RSABSEN.EOF Then
    MsgBox "ANDA BELUM ABSEN MASUK ..."
    'NIP.Enabled = True
    'NAMA.Enabled = True
    'NIP = ""
    'NIP.SetFocus
    JAMMASUK.SetFocus
    Exit Sub
Else
    Dim UPDATE As String
    UPDATE = "UPDATE ABSEN SET KELUAR='" & WAKTU & "'  WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "'"
    Conn.Execute UPDATE
    
    Dim UPDATE2 As String
    UPDATE2 = "UPDATE ABSEN SET LMKERJA='" & Format(RSABSEN!KELUAR - RSABSEN!MASUK, "HH:MM:SS") & "' WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "'"
    Conn.Execute UPDATE2
    
    Dim UPDATE1 As String
    'UPDATE1 = "UPDATE ABSEN SET LMLEMBUR='" & IIf(CDate(RSABSEN!LMKERJA) > CDate(NORMAL), Format(CDate(RSABSEN!LMKERJA) - CDate(NORMAL), "HH:MM:SS"), CDate("00:00:00")) & "' WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "'"
    UPDATE1 = "UPDATE ABSEN SET LMLEMBUR='" & IIf(CDate(RSABSEN!LMKERJA) > TimeValue("08:00:00"), Format(CDate(RSABSEN!LMKERJA) - TimeValue("08:00:00"), "HH:MM:SS"), CDate("00:00:00")) & "' WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "'"
    Conn.Execute UPDATE1
    MsgBox "OK...,DATA SUDAH DISIMPAN"
    NIP = ""
    NAMA = ""
    NIP.Enabled = True
    NAMA.Enabled = True
    NIP.SetFocus
End If

End Sub

Private Sub JAMMASUK_Click()
If NIP = "" Or NAMA = "" Then
    MsgBox "NIP DAN NAMA HARUS DIISI DULU"
    If NIP = "" Then
        NIP.SetFocus
        Exit Sub
    ElseIf NAMA = "" Then
        NAMA.SetFocus
        Exit Sub
    End If
End If

Call BukaDB
RSABSEN.Open "SELECT * FROM ABSEN WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "' AND MASUK<>NULL", Conn
If RSABSEN.EOF Then
    Dim SIMPAN As String
    SIMPAN = "INSERT INTO ABSEN (BULAN,NIP,MASUK) VALUES ('" & TANGGAL & "','" & NIP & "','" & WAKTU & "')"
    Conn.Execute SIMPAN
    MsgBox "DATA SUDAH DISIMPAN"
    NIP.Enabled = True
    NAMA.Enabled = True
    NIP = ""
    NAMA = ""
    NIP.SetFocus
Else
    MsgBox "ANDA SUDAH LOGIN SEBELUMNYA..."
End If

End Sub

Private Sub NAMA_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSPegawai.Open "SELECT * FROM PEGAWAI WHERE NIP='" & NIP & "' AND NAMA='" & NAMA & "'", Conn
    If RSPegawai.EOF Then
        MsgBox "NIP DAN NAMA TIDAK COCOK"
        NAMA = ""
        NAMA.SetFocus
        Exit Sub
    Else
        NAMA.Enabled = False
        MsgBox "OK, NIP DAN NAMA COCOK, SILAKAN PILIH ABSEN MASUK ATAU KELUAR"
    End If
End If
End Sub

Private Sub NIP_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Call BukaDB
    RSPegawai.Open "SELECT * FROM PEGAWAI WHERE NIP='" & NIP & "'", Conn
    If RSPegawai.EOF Then
        MsgBox "NIP TIDAK TERDAFTAR"
        NIP = ""
        NIP.SetFocus
        Exit Sub
    Else
        NIP.Enabled = False
        NAMA.SetFocus
    End If
End If
'        MsgBox "NAMA PEGAWAI : '" & RSPegawai!NAMA & "'"
'        RSABSEN.Open " select * from absen where nip='" & NIP & "' and cdate(bulan) = '" & TANGGAL & "'", Conn
'        If RSABSEN.EOF Then
'            JAMMASUK.Enabled = True
'            JAMMASUK.SetFocus
'            Exit Sub
'        Else
'            RSABSEN.Open "SELECT * FROM ABSEN WHERE NIP='" & NIP & "' AND CDATAE(BULAN) = '" & TANGGAL & "' AND "
'            JAMKELUAR.Enabled = True
'        End If
'    End If
'End If
End Sub

Private Sub Timer1_Timer()
WAKTU = Time$
End Sub
