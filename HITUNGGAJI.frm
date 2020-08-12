VERSION 5.00
Begin VB.Form HITUNGGAJI 
   Caption         =   "Perhitungan Gaji Karyawan"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox COMBO1 
      Height          =   345
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hitung Gaji"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label TAHUN 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TAHUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label BULAN 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BULAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label NOSLIP 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOSLIP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bulan / Tahun"
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1860
   End
End
Attribute VB_Name = "HITUNGGAJI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
NOSLIP = Format(Combo1, "YYMM") + "0000"
BULAN = Month(Combo1)
TAHUN = Year(Combo1)
End Sub

Private Sub Form_Load()
Call BukaDB
RSMASTER.Open "SELECT DISTINCT BULAN FROM MASTER", Conn
Combo1.Clear
Do While Not RSMASTER.EOF
    Combo1.AddItem RSMASTER!BULAN
    RSMASTER.MoveNext
Loop
End Sub

Private Sub Command1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
If Combo1 = "" Then
    MsgBox "PILIH DULU BULAN DAN TAHUNNYA"
    Combo1.SetFocus
    Exit Sub
End If

Call BukaDB
Dim RSHITUNG As New ADODB.Recordset
RSHITUNG.Open "SELECT PEGAWAI.NIP,NAMA,JABATAN.NMJABATAN AS JABATAN,PEGAWAI.GOL,STATUS,JMLANAK,GAPOK,JABATAN.TJJABATAN,IIF(PEGAWAI.STATUS='MENIKAH',TJSUAMIISTRI,0)  AS TJKELUARGA,IIF (STATUS='MENIKAH',TJANAK*JMLANAK,0)  AS TJANAK,UMAKAN*MASUK AS UANGMAKAN,MASTER.LEMBUR*GOLONGAN.LEMBUR AS UANGLEMBUR,ASKES, (GAPOK+TJJABATAN+TJKELUARGA+TJANAK+UANGMAKAN+ASKES+UANGLEMBUR) AS PENDAPATAN,  POTONGAN,PENDAPATAN-POTONGAN AS TOTALGAJI  FROM PEGAWAI,GOLONGAN,JABATAN,MASTER WHERE PEGAWAI.GOL=GOLONGAN.GOL AND PEGAWAI.KOJAB=JABATAN.KOJAB AND PEGAWAI.NIP=MASTER.NIP AND MONTH(BULAN)='" & Month(Combo1) & "' AND YEAR(BULAN)='" & Year(Combo1) & "'", Conn
If Not RSHITUNG.EOF Then
    RSGAJI.Open "SELECT * FROM GAJI WHERE MONTH(BULAN)='" & BULAN & "' AND YEAR(BULAN)='" & TAHUN & "'", Conn
    If RSGAJI.EOF Then
        Do While Not RSHITUNG.EOF
            NOSLIP = NOSLIP + 1
            Dim SIMPAN As String
            SIMPAN = "INSERT INTO GAJI(NOSLIP,BULAN,NIP,NAMA,JABATAN,GOL,STATUS,GAPOK,TJJABATAN,TJKELUARGA,TJANAK,UANGMAKAN,UANGLEMBUR,ASKES,PENDAPATAN,POTONGAN,TOTALGAJI) VALUES " & _
            "('" & NOSLIP & "','" & Combo1 & "','" & RSHITUNG!NIP & "','" & RSHITUNG!NAMA & "','" & RSHITUNG!JABATAN & "','" & RSHITUNG!GOL & "','" & RSHITUNG!Status & "','" & RSHITUNG!GAPOK & "','" & RSHITUNG!TJJABATAN & "', " & _
            "'" & RSHITUNG!TJKELUARGA & "','" & RSHITUNG!TJANAK & "','" & RSHITUNG!UANGMAKAN & "','" & RSHITUNG!UANGLEMBUR & "','" & RSHITUNG!ASKES & "','" & RSHITUNG!PENDAPATAN & "','" & RSHITUNG!POTONGAN & "','" & RSHITUNG!TotalGAJI & "')"
            Conn.Execute SIMPAN
            RSHITUNG.MoveNext
        Loop
    Else
        Do While Not RSHITUNG.EOF
            Dim EDIT As String
            EDIT = "UPDATE GAJI SET UANGMAKAN='" & RSHITUNG!UANGMAKAN & "',UANGLEMBUR='" & RSHITUNG!UANGLEMBUR & "',PENDAPATAN='" & RSHITUNG!PENDAPATAN & "',POTONGAN='" & RSHITUNG!POTONGAN & "',TOTALGAJI='" & RSHITUNG!TotalGAJI & "' WHERE NIP= '" & RSHITUNG!NIP & "' AND MONTH(BULAN)='" & Month(Combo1) & "' AND YEAR(BULAN)='" & Year(Combo1) & "'"
            Conn.Execute EDIT
            RSHITUNG.MoveNext
        Loop
  
    End If
        MsgBox "PERHITUNGAN GAJI BULAN '" & Month(Combo1) & "' TAHUN '" & Year(Combo1) & "' SUKSES"
End If
End Sub

