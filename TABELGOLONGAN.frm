VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TABELGOLONGAN 
   Caption         =   "TABEL GOLONGAN"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7020
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
      Left            =   360
      TabIndex        =   0
      Top             =   1560
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
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
      Left            =   3360
      TabIndex        =   2
      Top             =   1560
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
      Left            =   4920
      TabIndex        =   3
      Top             =   1560
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "TABELGOLONGAN.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
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
      ColumnCount     =   6
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "TJSUAMIISTRI"
         Caption         =   "TJSUAMIISTRI"
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
      BeginProperty Column02 
         DataField       =   "TJANAK"
         Caption         =   "TJANAK"
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
         DataField       =   "UMAKAN"
         Caption         =   "UMAKAN"
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
      BeginProperty Column04 
         DataField       =   "LEMBUR"
         Caption         =   "LEMBUR"
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
      BeginProperty Column05 
         DataField       =   "ASKES"
         Caption         =   "ASKES"
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
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
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
      Left            =   5160
      TabIndex        =   9
      Top             =   840
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
      Top             =   840
      Width           =   1500
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
      Left            =   5160
      TabIndex        =   7
      Top             =   480
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
      Top             =   480
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
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   1500
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   1920
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
      TabIndex        =   17
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Askes"
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
      Left            =   3600
      TabIndex        =   15
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Lembur"
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
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Uang Makan"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tunj. Anak"
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
      TabIndex        =   12
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tunj. Suami/Istri"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Golongan"
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
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "TABELGOLONGAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
Adodc1.RecordSource = "SELECT * FROM GOLONGAN ORDER BY GOL"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Sub Form_Load()
Call BukaDB
RSGOL.Open "SELECT DISTINCT GOL FROM GOLONGAN ORDER BY GOL", Conn
RSGOL.MoveFirst
Do While Not RSGOL.EOF
    Combo1.AddItem RSGOL!GOL
    RSGOL.MoveNext
Loop
KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSGOL.Open "Select * From GOLONGAN where GOL='" & Combo1 & "'", Conn
End Function

Private Sub KosongkanText()
    Combo1 = ""
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    
End Sub

Private Sub SiapIsi()
    Combo1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    Combo1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    Text5.Enabled = False
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
    With RSGOL
        If Not RSGOL.EOF Then
            Text1 = RSGOL!TJSUAMIISTRI
            Text2 = RSGOL!TJANAK
            Text3 = RSGOL!UMAKAN
            Text4 = RSGOL!LEMBUR
            Text5 = RSGOL!ASKES
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
        If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into GOLONGAN (GOL,TJSUAMIISTRI,TJANAK,UMAKAN,LEMBUR,ASKES) values ('" & Combo1 & "','" & Text1 & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "')"
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
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5 = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update GOLONGAN Set TJSUAMIISTRI= '" & Text1 & "', TJANAK='" & Text2 & "', UMAKAN='" & Text3 & "',LEMBUR='" & Text4 & "',ASKES='" & Text5 & "' where GOL='" & Combo1 & "'"
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
            If Not RSGOL.EOF Then
                TampilkanData
                MsgBox "Kode GOLONGAN Sudah Ada"
                KosongkanText
                Combo1.SetFocus
            Else
                Text1.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSGOL.EOF Then
                TampilkanData
                Combo1.Enabled = False
                Text1.SetFocus
            Else
                MsgBox "Kode GOLONGAN Tidak Ada"
                Combo1 = ""
                Combo1.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSGOL.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From GOLONGAN where GOL= '" & Combo1 & "'"
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
If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
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
    If Keyascii = 13 Then Text4.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TEXT4_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then Text5.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub TEXT5_Keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub


'Function CariGrid()
'    Call BukaDB
'    'mencari kode GOLONGAN tang ada dalam grid di kolom 0
'    RSGOL.Open "Select * From GOLONGAN where GOL='" & DataGrid1.Columns(0) & "'", Conn
'End Function
'
'Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Select Case KeyCode
'    'jika menekan enter setelah memilih data
'    Case vbKeyReturn
'        'jika cmdedit caption-nya simpan maka
'        If CmdEdit.Caption = "&Simpan" Then
'            'panggil prosedur SelectAllVisible1
'            Call SelectAllVisible1
'            Text1.SetFocus
'        'jika cmdhapus caption-nya hapus maka
'        ElseIf CmdHapus.Caption = "&Hapus" Then
'            'panggil prosedur SelectAllVisible2
'            Call SelectAllVisible2
'        End If
'    Case vbKeyEscape
'        KondisiAwal
'        CmdHapus.SetFocus
'End Select
'End Sub
'
'Sub SelectAllVisible1()
'    'jika COMBO1 tidak sama dengan isi grid kolom 0 maka
'    If Combo1 <> DataGrid1.Columns(0) Then
'        'ubah COMBO1 menjadi isi grid kolom 0 (GOL)
'        Combo1 = DataGrid1.Columns(0)
'        'panggil prosedur caridata
'        Call CariData
'        'COMBO1 (GOL) dinonaktifkan
'        Combo1.Enabled = False
'        'pindahkan isi grid kolom 0 ke COMBO1 dan seterusnya
'        Combo1 = DataGrid1.Columns(0)
'        Text1 = DataGrid1.Columns(1)
'        Text2 = DataGrid1.Columns(2)
'        Text3 = DataGrid1.Columns(3)
'        Text4 = DataGrid1.Columns(4)
'        Text5 = DataGrid1.Columns(5)
'        Text1.SetFocus
'    End If
'End Sub
'
'Sub SelectAllVisible2()
'    If Combo1 <> DataGrid1.Columns(0) Then
'        Combo1 = DataGrid1.Columns(0)
'        Call CariData
'        Combo1.Enabled = False
'        Combo1 = DataGrid1.Columns(0)
'        Text1 = DataGrid1.Columns(1)
'        Text2 = DataGrid1.Columns(2)
'        Text3 = DataGrid1.Columns(3)
'        Text4 = DataGrid1.Columns(4)
'        Text5 = DataGrid1.Columns(5)
'        'jika semua textbox telah terisi dan kode GOLONGAN ditemukan
'        'munculkan pesan penghapusan
'        Pesan = MsgBox("Yakin akan dihapus..?", vbYesNo, "Konfirmasi")
'        'jika dijawab YES
'        If Pesan = vbYes Then
'            'hapus data
'            Dim SQLHapus As String
'            SQLHapus = "Delete From GOLONGAN where GOL= '" & Combo1 & "'"
'            Conn.Execute SQLHapus
'            DataGrid1.Refresh
'            KondisiAwal
'            CmdHapus.SetFocus
'        Else
'            'jika dijawab NO kembali ke kondisi awal
'            KondisiAwal
'            CmdHapus.SetFocus
'        End If
'    End If
'End Sub
'
'Private Sub Text6_Change()
'Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
'Adodc1.RecordSource = "select * from GOLONGAN where GOL like '%" & Text6 & "%'"
'Adodc1.Refresh
'Set DataGrid1.DataSource = Adodc1
'DataGrid1.Refresh
'End Sub
'
'Private Sub COMBO1_Change()
'Adodc1.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGAJI.mdb"
'Adodc1.RecordSource = "select * from GOLONGAN where GOL like '%" & Combo1 & "%'"
'Adodc1.Refresh
'Set DataGrid1.DataSource = Adodc1
'DataGrid1.Refresh
'End Sub
'
'
'Private Sub TEXT6_Keypress(KeyAscii As Integer)
'If KeyAscii = 13 Then DataGrid1.SetFocus
'End Sub
'
