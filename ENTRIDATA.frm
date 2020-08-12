VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ENTRIDATA 
   Caption         =   "ENTRI DATA UNTUK HITUNG GAJI"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
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
   ScaleHeight     =   5325
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Format          =   111804417
      CurrentDate     =   39821
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mulai Entri"
      Height          =   350
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "ENTRIDATA.frx":0000
      Height          =   4035
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7117
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "NO"
         Caption         =   "NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NIP"
         Caption         =   "NIP"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "NAMA"
         Caption         =   "NAMA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "JABATAN"
         Caption         =   "JABATAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "MASUK"
         Caption         =   "MASUK"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "SAKIT"
         Caption         =   "SAKIT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "IZIN"
         Caption         =   "IZIN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "ALPA"
         Caption         =   "ALPA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "LEMBUR"
         Caption         =   "LEMBUR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "POTONGAN"
         Caption         =   "POTONGAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2505,26
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   1200
      Top             =   4800
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pilih Bulan Dan Tahun"
      Height          =   225
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " Jumlah Data"
      Height          =   225
      Left            =   9600
      TabIndex        =   4
      Top             =   4800
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   225
      Left            =   10680
      TabIndex        =   3
      Top             =   4800
      Width           =   510
   End
End
Attribute VB_Name = "ENTRIDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGaji.mdb"
Adodc1.RecordSource = "select * from TEMPORER ORDER BY NIP"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh


End Sub

Private Sub Form_Load()
Call BukaDB
Form_Activate

DTPicker1.Value = Date

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

Private Sub Combo1_Click()
Call BukaDB
RSMASTER.Open "SELECT * FROM MASTER WHERE MONTH(BULAN)='" & Month(Combo1) & "' AND YEAR(BULAN)='" & Year(Combo1) & "'", Conn
Combo1 = RSMASTER!BULAN
If Not RSMASTER.EOF Then
    RSMASTER.MoveFirst
    Do While Not RSMASTER.EOF
        Dim UPDATE As String
        UPDATE = "UPDATE TEMPORER SET MASUK='" & RSMASTER!masuk & "',SAKIT='" & RSMASTER!SAKIT & "',IZIN='" & RSMASTER!IZIN & "',ALPA='" & RSMASTER!ALPA & "',LEMBUR='" & RSMASTER!LEMBUR & "',POTONGAN='" & RSMASTER!POTONGAN & "' WHERE NIP='" & RSMASTER!NIP & "'" ' AND MONTH(BULAN)='" & Month(Combo1) & "' AND YEAR(BULAN)='" & Year(Combo1) & "'"
        Conn.Execute UPDATE
        RSMASTER.MoveNext
    Loop
    Call TAMPILKAN
    MsgBox "DATA BULAN '" & Month(Combo1) & "' TAHUN '" & Year(Combo1) & "' SUDAH ADA, SILAKAN EDIT, KEMUDIAN SIMPAN..."
End If

End Sub

Sub TAMPILKAN()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGaji.mdb"
Adodc1.RecordSource = "select * from TEMPORER ORDER BY NIP"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh

Label3 = Adodc1.Recordset.RecordCount
DataGrid1.Col = 3
Adodc1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command1_Click()
Call BukaDB
RSMASTER.Open "SELECT * FROM MASTER WHERE MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'", Conn
If Not RSMASTER.EOF Then
    RSMASTER.MoveFirst
    Do While Not RSMASTER.EOF
        Dim UPDATE As String
        UPDATE = "UPDATE TEMPORER SET MASUK='" & RSMASTER!masuk & "',SAKIT='" & RSMASTER!SAKIT & "',IZIN='" & RSMASTER!IZIN & "',ALPA='" & RSMASTER!ALPA & "',LEMBUR='" & RSMASTER!LEMBUR & "',POTONGAN='" & RSMASTER!POTONGAN & "' " & _
        "WHERE NIP='" & RSMASTER!NIP & "'" '  AND MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "' "
        Conn.Execute UPDATE
        RSMASTER.MoveNext
    Loop
    Call TAMPILKAN
    MsgBox "DATA BULAN '" & Month(DTPicker1) & "' TAHUN '" & Year(DTPicker1) & "' SUDAH ADA, SILAKAN EDIT KEMUDIAN SIMPAN..."
    
Else
    Call BukaDB
    Dim RSTEMPORER As New ADODB.Recordset
    RSTEMPORER.Open "TEMPORER", Conn
    RSTEMPORER.MoveFirst
    Do While Not RSTEMPORER.EOF
        Dim Kosongkan As String
        Kosongkan = "UPDATE TEMPORER SET MASUK=0,SAKIT=0,IZIN=0,ALPA=0,LEMBUR=0,POTONGAN=0 "
        Conn.Execute Kosongkan
        RSTEMPORER.MoveNext
    Loop
    Call TAMPILKAN
    MsgBox "DATA BULAN '" & Month(DTPicker1) & "' TAHUN '" & Year(DTPicker1) & "' MASIH KOSONG, SILAKAN ENTRI..."
End If
End Sub

Private Sub Command1_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MoveFirst
Do While Not Adodc1.Recordset.EOF
    Call BukaDB
    RSMASTER.Open "SELECT * FROM MASTER WHERE NIP='" & Adodc1.Recordset!NIP & "' and MONTH(bulan)= '" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'", Conn
    If Not RSMASTER.EOF Then
        Dim UPDATE As String
        UPDATE = "UPDATE MASTER SET MASUK='" & Adodc1.Recordset!masuk & "',SAKIT='" & Adodc1.Recordset!SAKIT & "',IZIN='" & Adodc1.Recordset!IZIN & "',ALPA='" & Adodc1.Recordset!ALPA & "',LEMBUR='" & Adodc1.Recordset!LEMBUR & "',POTONGAN ='" & Adodc1.Recordset!POTONGAN & "' WHERE NIP='" & Adodc1.Recordset!NIP & "' AND MONTH(BULAN)='" & Month(DTPicker1) & "' AND YEAR(BULAN)='" & Year(DTPicker1) & "'"
        Conn.Execute UPDATE
        Adodc1.Recordset.MoveNext
    Else
        Dim SIMPAN As String
        SIMPAN = "INSERT INTO MASTER (BULAN,NIP,MASUK,SAKIT,IZIN,ALPA,LEMBUR,POTONGAN) VALUES " & _
        "('" & DTPicker1 & "','" & Adodc1.Recordset!NIP & "','" & Adodc1.Recordset!masuk & "','" & Adodc1.Recordset!SAKIT & "','" & Adodc1.Recordset!IZIN & "','" & Adodc1.Recordset!ALPA & "','" & Adodc1.Recordset!LEMBUR & "','" & Adodc1.Recordset!POTONGAN & "')"
        Conn.Execute SIMPAN
        Adodc1.Recordset.MoveNext
    End If
Loop
Form_Activate
End Sub

Private Sub DataGrid1_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
'MASUK
If DataGrid1.Col = 4 Then
    Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 4
'SAKIT
ElseIf DataGrid1.Col = 5 Then
   Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 5
'IZIN
ElseIf DataGrid1.Col = 6 Then
    Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 6
'ALPA
ElseIf DataGrid1.Col = 7 Then
   Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 7
'LEMBUR
ElseIf DataGrid1.Col = 8 Then
    Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 8
'POTONGAN
ElseIf DataGrid1.Col = 9 Then
    Adodc1.Recordset.UPDATE
    Adodc1.Recordset.MoveNext
    DataGrid1.Col = 9
End If
End Sub



