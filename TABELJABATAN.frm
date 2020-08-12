VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TABELJABATAN 
   Caption         =   "TABEL JABATAN"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3120
      TabIndex        =   2
      Top             =   1800
      Width           =   1500
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   1500
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   4620
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   4620
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "TABELJABATAN.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4471
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
      ColumnCount     =   4
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "NMJABATAN"
         Caption         =   "NMJABATAN"
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
         DataField       =   "GAPOK"
         Caption         =   "GAPOK"
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
      BeginProperty Column03 
         DataField       =   "TJJABATAN"
         Caption         =   "TJJABATAN"
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
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cari Jabatan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tunj. Jabatan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   13
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Gaji Pokok"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama Jabatan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Kode Jabatan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "TABELJABATAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
Adodc1.RecordSource = "SELECT * FROM JABATAN ORDER BY KOJAB"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
Call BukaDB
RSJabatan.Open "SELECT DISTINCT KOJAB FROM JABATAN ORDER BY KOJAB", Conn
RSJabatan.MoveFirst
Do While Not RSJabatan.EOF
    Combo1.AddItem RSJabatan!KOJAB
    RSJabatan.MoveNext
Loop
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSJabatan.Open "Select * From JABATAN where KOJAB='" & Combo1 & "'", Conn
End Function

Private Sub KosongkanText()
    Combo1 = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
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
    With RSJabatan
        If Not RSJabatan.EOF Then
            Text1 = RSJabatan!NMJABATAN
            Text2 = RSJabatan!GAPOK
            Text3 = RSJabatan!TJJABATAN
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
        Combo1.SetFocus
    Else
        If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into JABATAN (KOJAB,NMJABATAN,GAPOK,TJJABATAN) values ('" & Combo1 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "')"
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
        Combo1.SetFocus
    Else
        If Text1 = "" Or Text2 = "" Or Text3 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update JABATAN Set NMJABATAN= '" & Text1 & "', GAPOK='" & Text2 & "', TJJABATAN='" & Text3 & "' where KOJAB='" & Combo1 & "'"
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
        Combo1.SetFocus
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

Private Sub Combo1_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(Combo1) < 1 Then
        MsgBox "Kode Harus 1 Digit"
        Combo1.SetFocus
    Else
        Text1.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSJabatan.EOF Then
                TampilkanData
                MsgBox "Kode JABATAN Sudah Ada"
                KosongkanText
                Combo1.SetFocus
            Else
                Text1.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSJabatan.EOF Then
                TampilkanData
                Combo1.Enabled = False
                Text1.SetFocus
            Else
                MsgBox "Kode JABATAN Tidak Ada"
                Combo1 = ""
                Combo1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSJabatan.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From JABATAN where KOJAB= '" & Combo1 & "'"
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
                Combo1.SetFocus
            End If
    End If
End If
'If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub Text1_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Text2.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TEXT2_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then Text3.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
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
    'mencari kode JABATAN tang ada dalam grid di kolom 0
    RSJabatan.Open "Select * From JABATAN where KOJAB='" & DataGrid1.Columns(0) & "'", Conn
End Function

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    'jika menekan enter setelah memilih data
    Case vbKeyReturn
        'jika cmdedit caption-nya simpan maka
        If CmdEdit.Caption = "&Simpan" Then
            'panggil prosedur SelectAllVisible1
            Call SelectAllVisible1
            Text1.SetFocus
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
    'jika COMBO1 tidak sama dengan isi grid kolom 0 maka
    If Combo1 <> DataGrid1.Columns(0) Then
        'ubah COMBO1 menjadi isi grid kolom 0 (KOJAB)
        Combo1 = DataGrid1.Columns(0)
        'panggil prosedur caridata
        Call CariData
        'COMBO1 (KOJAB) dinonaktifkan
        Combo1.Enabled = False
        'pindahkan isi grid kolom 0 ke COMBO1 dan seterusnya
        Combo1 = DataGrid1.Columns(0)
        Text1 = DataGrid1.Columns(1)
        Text2 = DataGrid1.Columns(2)
        Text3 = DataGrid1.Columns(3)
        Text1.SetFocus
    End If
End Sub

Sub SelectAllVisible2()
    If Combo1 <> DataGrid1.Columns(0) Then
        Combo1 = DataGrid1.Columns(0)
        Call CariData
        Combo1.Enabled = False
        Combo1 = DataGrid1.Columns(0)
        Text1 = DataGrid1.Columns(1)
        Text2 = DataGrid1.Columns(2)
        Text3 = DataGrid1.Columns(3)
        'jika semua textbox telah terisi dan kode JABATAN ditemukan
        'munculkan pesan penghapusan
        Pesan = MsgBox("Yakin akan dihapus..?", vbYesNo, "Konfirmasi")
        'jika dijawab YES
        If Pesan = vbYes Then
            'hapus data
            Dim SQLHapus As String
            SQLHapus = "Delete From JABATAN where KOJAB= '" & Combo1 & "'"
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
Adodc1.RecordSource = "select * from JABATAN where NMJABATAN like '%" & Text4 & "%'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

'Private Sub COMBO1_Change()
'Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
'Adodc1.RecordSource = "select * from JABATAN where KOJAB like '%" & Combo1 & "%'"
'Adodc1.Refresh
'Set DataGrid1.DataSource = Adodc1
'DataGrid1.Refresh
'End Sub


Private Sub TEXT4_Keypress(Keyascii As Integer)
If Keyascii = 13 Then DataGrid1.SetFocus
End Sub


