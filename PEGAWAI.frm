VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PEGAWAI 
   Caption         =   "DATA PEGAWAI"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
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
   ScaleHeight     =   5175
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   350
      Left            =   1560
      TabIndex        =   10
      Top             =   1680
      Width           =   5000
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   400
      Left            =   4920
      TabIndex        =   3
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   400
      Left            =   3360
      TabIndex        =   2
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   400
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PEGAWAI.frx":0000
      Height          =   1935
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
         DataField       =   "KOJAB"
         Caption         =   "KOJAB"
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
         DataField       =   "GOL"
         Caption         =   "GOL"
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
         DataField       =   "STATUS"
         Caption         =   "STATUS"
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
         DataField       =   "JMLANAK"
         Caption         =   "JMLANAK"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4440
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
   Begin VB.ComboBox Combo3 
      Height          =   345
      Left            =   4560
      TabIndex        =   8
      Top             =   840
      Width           =   2000
   End
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Height          =   350
      Left            =   4560
      TabIndex        =   9
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   350
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      Height          =   350
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   6495
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cari Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah Anak"
      Height          =   345
      Left            =   3120
      TabIndex        =   16
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Status"
      Height          =   345
      Left            =   3120
      TabIndex        =   15
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Golongan"
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Jabatan"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Karyawan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1400
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIP"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1400
   End
End
Attribute VB_Name = "PEGAWAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
Adodc1.RecordSource = "SELECT * FROM PEGAWAI ORDER BY NIP"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
Text1.MaxLength = 4
Text2.MaxLength = 30
Text3.MaxLength = 2

Call BukaDB
RSPegawai.Open "SELECT DISTINCT KOJAB FROM PEGAWAI ORDER BY KOJAB", Conn
RSPegawai.MoveFirst
Do While Not RSPegawai.EOF
    Combo1.AddItem RSPegawai!KOJAB
    RSPegawai.MoveNext
Loop

For I = 1 To 6
    Combo2.AddItem I
Next I

Combo3.AddItem "MENIKAH"
Combo3.AddItem "TIDAK MENIKAH"

KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSPegawai.Open "Select * From PEGAWAI where NIP='" & Text1 & "'", Conn
End Function

Private Sub KosongkanText()
    Text1 = ""
    Text2 = ""
    Combo1 = ""
    Combo2 = ""
    Combo3 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Text1.Enabled = True
    Text2.Enabled = True
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Text1.Enabled = False
    Text2.Enabled = False
    Combo1.Enabled = False
    Combo2.Enabled = False
    Combo3.Enabled = False
    Text3.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSPegawai
        If Not RSPegawai.EOF Then
            Text2 = RSPegawai!NAMA
            Combo1 = RSPegawai!KOJAB
            Combo2 = RSPegawai!GOL
            Combo3 = RSPegawai!Status
            Text3 = RSPegawai!JNLANAK
        End If
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Text1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Text3 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into PEGAWAI (NIP,NAMA,KOJAB,GOL,STATUS,JMLANAK) values ('" & Text1 & "','" & Text2 & "','" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & Text3 & "')"
            Conn.Execute SQLTambah
            Call KondisiAwal
            Form_Activate
        End If
    End If
End Sub

Private Sub CmdEdit_Click()

    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        Text4.SetFocus
    Else
        If Text2 = "" Or Combo1 = "" Or Combo2 = "" Or Combo3 = "" Or Text3 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update PEGAWAI Set NAMA= '" & Text2 & "', KOJAB='" & Combo1 & "', GOL='" & Combo2 & "',STATUS='" & Combo3 & "',JMLANAK='" & Text3 & "' where NIP='" & Text1 & "'"
            Conn.Execute SQLEdit
            Call KondisiAwal
            Form_Activate
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        Text4.SetFocus
    End If
End Sub

Private Sub CmdTutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Text1) < 4 Then
        MsgBox "Kode Harus 4 Digit"
        Text1.SetFocus
    Else
        Text2.SetFocus
    End If

    If CmdInput.Caption = "SIMPAN" Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                MsgBox "Kode PEGAWAI Sudah Ada"
                KosongkanText
                Text1.SetFocus
            Else
                Text2.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "SIMPAN" Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                Text1.Enabled = False
                Text2.SetFocus
            Else
                MsgBox "Kode PEGAWAI Tidak Ada"
                Text1 = ""
                Text1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSPegawai.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From PEGAWAI where NIP= '" & Text1 & "'"
                    Conn.Execute SQLHapus
                    Call KondisiAwal
                    Form_Activate
                    CmdHapus.SetFocus
                Else
                    Call KondisiAwal
                    Form_Activate
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                Text1.SetFocus
            End If
    End If
End If
End Sub

Private Sub TEXT2_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Combo1.SetFocus
End Sub

Private Sub Combo1_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Combo2.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub COMBO2_Keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Combo3.SetFocus
End Sub

Private Sub COMBO3_Keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then Text3.SetFocus
End Sub

Private Sub TEXT3_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


Function CariGrid()
    Call BukaDB
    'mencari kode PEGAWAI tang ada dalam grid di kolom 0
    RSPegawai.Open "Select * From PEGAWAI where NIP='" & DataGrid1.Columns(0) & "'", Conn
End Function

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    'jika menekan enter setelah memilih data
    Case vbKeyReturn
        'jika cmdedit caption-nya simpan maka
        If CmdEdit.Caption = "&Simpan" Then
            'panggil prosedur SelectAllVisible1
            Call SelectAllVisible1
            Text2.SetFocus
        'jika cmdhapus caption-nya hapus maka
        ElseIf CmdHapus.Caption = "&Hapus" Then
            'panggil prosedur SelectAllVisible2
            Call SelectAllVisible2
        End If
    Case vbKeyEscape
        KondisiAwal
        CmdHapus.SetFocus
End Select
End Sub

Sub SelectAllVisible1()
    'jika text1 tidak sama dengan isi grid kolom 0 maka
    If Text1 <> DataGrid1.Columns(0) Then
        'ubah text1 menjadi isi grid kolom 0 (NIP)
        Text1 = DataGrid1.Columns(0)
        'panggil prosedur caridata
        Call CariData
        'text1 (NIP) dinonaktifkan
        Text1.Enabled = False
        'pindahkan isi grid kolom 0 ke text1 dan seterusnya
        Text1 = DataGrid1.Columns(0)
        Text2 = DataGrid1.Columns(1)
        Combo1 = DataGrid1.Columns(2)
        Combo2 = DataGrid1.Columns(3)
        Combo3 = DataGrid1.Columns(4)
        Text3 = DataGrid1.Columns(5)
        Text2.SetFocus
    End If
End Sub

Sub SelectAllVisible2()
    If Text1 <> DataGrid1.Columns(0) Then
        Text1 = DataGrid1.Columns(0)
        Call CariData
        Text1.Enabled = False
        Text1 = DataGrid1.Columns(0)
        Text2 = DataGrid1.Columns(1)
        Combo1 = DataGrid1.Columns(2)
        Combo2 = DataGrid1.Columns(3)
        Combo3 = DataGrid1.Columns(4)
        Text3 = DataGrid1.Columns(5)
        'jika semua textbox telah terisi dan kode PEGAWAI ditemukan
        'munculkan pesan penghapusan
        Pesan = MsgBox("Yakin akan dihapus..?", vbYesNo, "Konfirmasi")
        'jika dijawab YES
        If Pesan = vbYes Then
            'hapus data
            Dim SQLHapus As String
            SQLHapus = "Delete From PEGAWAI where NIP= '" & Text1 & "'"
            Conn.Execute SQLHapus
            DataGrid1.Refresh
            KondisiAwal
            CmdHapus.SetFocus
        Else
            'jika dijawab NO kembali ke kondisi awal
            KondisiAwal
            CmdHapus.SetFocus
        End If
    End If
End Sub

Private Sub Text4_Change()
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
Adodc1.RecordSource = "select * from PEGAWAI where NAMA like '%" & Text4 & "%'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Text1_Change()
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
Adodc1.RecordSource = "select * from PEGAWAI where NIP like '%" & Text1 & "%'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub


Private Sub TEXT4_Keypress(Keyascii As Integer)
If Keyascii = 13 Then DataGrid1.SetFocus
End Sub
