VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ABSEN1 
   Caption         =   "ABSEN DERET HARI"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
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
   ScaleHeight     =   8985
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   92340225
      CurrentDate     =   39830
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   400
      Left            =   120
      TabIndex        =   1
      Top             =   8400
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MULAI ENTRI"
      Height          =   350
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ABSEN1.frx":0000
      Height          =   7275
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12832
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "TANGGAL"
         Caption         =   "TANGGAL"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MASUK"
         Caption         =   "MASUK"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "KELUAR"
         Caption         =   "KELUAR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LMKERJA"
         Caption         =   "LMKERJA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "LMLEMBUR"
         Caption         =   "LMLEMBUR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   4
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   1200
      Top             =   8400
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   4740
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NAMA "
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label NAMABULAN 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   225
      Left            =   5160
      TabIndex        =   5
      Top             =   8400
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " Jumlah Data"
      Height          =   225
      Left            =   4080
      TabIndex        =   4
      Top             =   8400
      Width           =   1035
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NIP"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "ABSEN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1 = "1" Then NAMABULAN = "JANUARI"
If Combo1 = "2" Then NAMABULAN = "FEBRUARI"
If Combo1 = "3" Then NAMABULAN = "MARET"
If Combo1 = "4" Then NAMABULAN = "APRIL"
If Combo1 = "5" Then NAMABULAN = "MEI"
If Combo1 = "6" Then NAMABULAN = "JUNI"
If Combo1 = "7" Then NAMABULAN = "JULI"
If Combo1 = "8" Then NAMABULAN = "AGUSTUS"
If Combo1 = "9" Then NAMABULAN = "SEPTEMBER"
If Combo1 = "10" Then NAMABULAN = "OKTOBER"
If Combo1 = "11" Then NAMABULAN = "NOVEMBER"
If Combo1 = "12" Then NAMABULAN = "DESEMBER"


End Sub

Private Sub Combo1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command1_Click()
If Day(DTPicker1) <> 1 Then
    MsgBox "HARUS DIMULAI DARI TANGGAL 1"
    Exit Sub
    DTPicker1.SetFocus
End If

Dim QQQ As String
QQQ = "DELETE * FROM TRABSEN1"
Conn.Execute QQQ

For I = 0 To 30
    Dim AAA As String
    AAA = "INSERT INTO TRABSEN1(TANGGAL) VALUES ('" & DTPicker1 + I & "')"
    Conn.Execute AAA
Next I
Form_Activate
DataGrid1.Enabled = True
DataGrid1.Col = 2
End Sub

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGaji.mdb"
Adodc1.RecordSource = "select * from TRABSEN1 "
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh


End Sub

Private Sub Form_Load()
DataGrid1.Enabled = False
'For I = 1 To 12
'    COMBO1.AddItem I
'Next I
Call BukaDB
Form_Activate

'DTPicker1.Value = Date

RSTEMPORER.Open "TEMPORER", Conn
If RSTEMPORER.EOF Then
    Dim RS As New ADODB.Recordset
    RS.Open "SELECT PEGAWAI.NIP,PEGAWAI.NAMA,JABATAN.NMJABATAN  FROM PEGAWAI,JABATAN WHERE PEGAWAI.KOJAB=JABATAN.KOJAB ORDER BY NIP ", Conn
    RS.MoveFirst
    Dim NO As Byte
    NO = 0
    Do While Not RS.EOF
        NO = NO + 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!NO = NO
        Adodc1.Recordset!NIP = RS!NIP
        Adodc1.Recordset!NAMA = RS!NAMA
        Adodc1.Recordset!JABATAN = RS!NMJABATAN
        Adodc1.Recordset!masuk = 0
        Adodc1.Recordset!SAKIT = 0
        Adodc1.Recordset!IZIN = 0
        Adodc1.Recordset!ALPA = 0
        Adodc1.Recordset!LEMBUR = 0
        Adodc1.Recordset!POTONGAN = 0
        Adodc1.Recordset.UPDATE
        RS.MoveNext
    Loop
    Adodc1.Refresh
    DataGrid1.Refresh
    'Form_Activate
    Call TAMPILKAN
Else
    RSTEMPORER.MoveFirst
    Do While Not RSTEMPORER.EOF
        Dim NOLKAN As String
        NOLKAN = "UPDATE TEMPORER SET MASUK=0,SAKIT=0,IZIN=0,ALPA=0,LEMBUR=0,POTONGAN=0"
        Conn.Execute NOLKAN
        RSTEMPORER.MoveNext
    Loop
    Adodc1.Refresh
    DataGrid1.Refresh
    Call TAMPILKAN
End If

End Sub

'Private Sub Combo1_Click()
'Call BukaDB
'RSMASTER.Open "SELECT * FROM MASTER WHERE MONTH(BULAN)='" & Month(COMBO1) & "' AND YEAR(BULAN)='" & Year(COMBO1) & "'", Conn
'COMBO1 = RSMASTER!BULAN
'If Not RSMASTER.EOF Then
'    RSMASTER.MoveFirst
'    Do While Not RSMASTER.EOF
'        Dim UPDATE As String
'        UPDATE = "UPDATE TEMPORER SET MASUK='" & RSMASTER!MASUK & "',SAKIT='" & RSMASTER!SAKIT & "',IZIN='" & RSMASTER!IZIN & "',ALPA='" & RSMASTER!ALPA & "',LEMBUR='" & RSMASTER!LEMBUR & "',POTONGAN='" & RSMASTER!POTONGAN & "' WHERE NIP='" & RSMASTER!NIP & "'" ' AND MONTH(BULAN)='" & Month(Combo1) & "' AND YEAR(BULAN)='" & Year(Combo1) & "'"
'        Conn.Execute UPDATE
'        RSMASTER.MoveNext
'    Loop
'    Call TAMPILKAN
'    MsgBox "DATA BULAN '" & Month(COMBO1) & "' TAHUN '" & Year(COMBO1) & "' SUDAH ADA, SILAKAN EDIT, KEMUDIAN SIMPAN..."
'End If
'
'End Sub

Sub TAMPILKAN()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGaji.mdb"
Adodc1.RecordSource = "select * from TRABSEN1"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Label3 = Adodc1.Recordset.RecordCount
DataGrid1.Col = 3
Adodc1.Refresh
DataGrid1.Refresh
End Sub

'Private Sub Command1_Click()
'Call BukaDB
'RSMASTER.Open "SELECT * FROM MASTER WHERE MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'", Conn
'If Not RSMASTER.EOF Then
'    RSMASTER.MoveFirst
'    Do While Not RSMASTER.EOF
'        Dim UPDATE As String
'        UPDATE = "UPDATE TEMPORER SET MASUK='" & RSMASTER!MASUK & "',SAKIT='" & RSMASTER!SAKIT & "',IZIN='" & RSMASTER!IZIN & "',ALPA='" & RSMASTER!ALPA & "',LEMBUR='" & RSMASTER!LEMBUR & "',POTONGAN='" & RSMASTER!POTONGAN & "' " & _
'        "WHERE NIP='" & RSMASTER!NIP & "'" '  AND MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "' "
'        Conn.Execute UPDATE
'        RSMASTER.MoveNext
'    Loop
'    Call TAMPILKAN
'    MsgBox "DATA BULAN '" & Month(DTPicker1) & "' TAHUN '" & Year(DTPicker1) & "' SUDAH ADA, SILAKAN EDIT KEMUDIAN SIMPAN..."
'
'Else
'    Call BukaDB
'    Dim RSTEMPORER As New ADODB.Recordset
'    RSTEMPORER.Open "TEMPORER", Conn
'    RSTEMPORER.MoveFirst
'    Do While Not RSTEMPORER.EOF
'        Dim Kosongkan As String
'        Kosongkan = "UPDATE TEMPORER SET MASUK=0,SAKIT=0,IZIN=0,ALPA=0,LEMBUR=0,POTONGAN=0 "
'        Conn.Execute Kosongkan
'        RSTEMPORER.MoveNext
'    Loop
'    Call TAMPILKAN
'    MsgBox "DATA BULAN '" & Month(DTPicker1) & "' TAHUN '" & Year(DTPicker1) & "' MASIH KOSONG, SILAKAN ENTRI..."
'End If
'End Sub

Private Sub Command1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!masuk <> vbNullString
    Call BukaDB
    RSMASTER.Open "SELECT * FROM absen WHERE NIP='" & Text1 & "' and MONTH(bulan)= '" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'", Conn
    If Not RSMASTER.EOF Then
        Dim UPDATE As String
        UPDATE = "UPDATE absen SET MASUK='" & Adodc1.Recordset!masuk & "',keluar ='" & Adodc1.Recordset!keluar & "',lmkerja='" & Adodc1.Recordset!lmkerja & "',lmlembur='" & Adodc1.Recordset!lmlembur & "' WHERE NIP='" & Text1 & "' AND MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'"
        Conn.Execute UPDATE
        Adodc1.Recordset.MoveNext
    Else
        Dim SIMPAN As String
        SIMPAN = "INSERT INTO absen (BULAN,NIP,MASUK,keluar,lmkerja,lmlembur) VALUES " & _
        "('" & DTPicker1 & "','" & Text1 & "','" & Adodc1.Recordset!masuk & "','" & Adodc1.Recordset!keluar & "','" & Adodc1.Recordset!lmkerja & "','" & Adodc1.Recordset!lmlembur & "')"
        Conn.Execute SIMPAN
        Adodc1.Recordset.MoveNext
    End If
Loop
Form_Activate
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
'MASUK
If DataGrid1.Col = 1 Then
    Adodc1.Recordset.UPDATE
    DataGrid1.Col = 2
    Exit Sub
End If

'KELUAR
If DataGrid1.Col = 2 Then
    Adodc1.Recordset!lmkerja = Format(Adodc1.Recordset!keluar - Adodc1.Recordset!masuk, "HH:MM:SS")
    Adodc1.Recordset!lmlembur = IIf(Adodc1.Recordset!lmkerja > TimeValue("08:00:00"), Adodc1.Recordset!keluar - -TimeValue("08:00:00"), CDate("00:00:00"))
    Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 1
End If


'UPDATE2 = "UPDATE ABSEN SET LMKERJA='" & Format(RSABSEN!KELUAR - RSABSEN!MASUK, "HH:MM:SS") & "' WHERE NIP='" & NIP & "' AND CDATE(BULAN)='" & TANGGAL & "'"
End Sub





Private Sub Text1_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then
    Call BukaDB
    RSPegawai.Open "SELECT * FROM PEGAWAI WHERE NIP='" & Text1 & "'", Conn
    If RSPegawai.EOF Then
        MsgBox " NIP TIDAK TERDAFTAR"
        Text1.SetFocus
        Exit Sub
    Else
        Label5 = RSPegawai!NAMA
        DTPicker1.SetFocus
    End If
End If
End Sub
