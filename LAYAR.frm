VERSION 5.00
Begin VB.Form LAYAR 
   BackColor       =   &H80000009&
   Caption         =   "ESC = TUTUP ** ENTER = PRINT"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
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
   ScaleHeight     =   6510
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "LAYAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Pesan = MsgBox("PRINTER SUDAH SIAP", vbYesNo)
    If Pesan = vbYes Then
        Call CetakGAJI
    End If
End If
End Sub



Function CetakGAJI()
Call BukaDB
Dim RSGAJI As New ADODB.Recordset
'RSGAJI.Open "SELECT PEGAWAI.NIP,NAMA,JABATAN.NMJABATAN AS JABATAN,PEGAWAI.GOL,STATUS,JMLANAK,MASUK,JAMLEMBUR,GAPOK,JABATAN.TJJABATAN,IIF(PEGAWAI.STATUS='MENIKAH',TJSUAMIISTRI,0)  AS TJKELUARGA,IIF (STATUS='MENIKAH',TJANAK*JMLANAK,0)  AS TJANAK,UMAKAN*MASUK AS UANGMAKAN,JAMLEMBUR*LEMBUR AS UANGLEMBUR,ASKES, (GAPOK+TJJABATAN+TJKELUARGA+TJANAK+UANGMAKAN+ASKES+UANGLEMBUR) AS PENDAPATAN,  POTONGAN,PENDAPATAN-POTONGAN AS TOTALGAJI  FROM PEGAWAI,GOLONGAN,JABATAN,LEMBUR,POTONGAN,ABSEN WHERE PEGAWAI.GOL=GOLONGAN.GOL AND PEGAWAI.KOJAB=JABATAN.KOJAB AND PEGAWAI.NIP=LEMBUR.NIP AND PEGAWAI.NIP=POTONGAN.NIP AND PEGAWAI.NIP=ABSEN.NIP AND PEGAWAI.NIP='" & SLIPGAJI.Combo2 & "'", Conn
RSGAJI.Open "SELECT PEGAWAI.NIP,NAMA,JABATAN.NMJABATAN AS JABATAN,PEGAWAI.GOL,STATUS,JMLANAK,MASUK,MASTER.LEMBUR,GAPOK,JABATAN.TJJABATAN,IIF(PEGAWAI.STATUS='MENIKAH',TJSUAMIISTRI,0)  AS TJKELUARGA,IIF (STATUS='MENIKAH',TJANAK*JMLANAK,0)  AS TJANAK,UMAKAN*MASUK AS UANGMAKAN,MASTER.LEMBUR*GOLONGAN.LEMBUR AS UANGLEMBUR,ASKES, (GAPOK+TJJABATAN+TJKELUARGA+TJANAK+UANGMAKAN+ASKES+UANGLEMBUR) AS PENDAPATAN,  POTONGAN,PENDAPATAN-POTONGAN AS TOTALGAJI  FROM PEGAWAI,GOLONGAN,JABATAN,MASTER WHERE PEGAWAI.GOL=GOLONGAN.GOL AND PEGAWAI.KOJAB=JABATAN.KOJAB AND PEGAWAI.NIP=MASTER.NIP AND PEGAWAI.NIP='" & SLIPGAJI.Combo2 & "'", Conn

Dim MGrs As String
MGrs = String$(40, "-")
Printer.Font = "Courier New"
Printer.Print
Printer.Print
'PRINTER.Print Tab(5); MGrs
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.Print Tab(5); "TANGGAL                :   "; Format(Date, "DD-MMM-YYYY")
Printer.Print
Printer.Print Tab(5); "NIP                    :   "; RSGAJI!NIP
Printer.Print Tab(5); "NAMA PEGAWAI           :   "; RSGAJI!NAMA
Printer.Print Tab(5); "JABATAN                :   "; RSGAJI!JABATAN
Printer.Print Tab(5); "GOLONGAN               :   "; RSGAJI!GOL
Printer.Print Tab(5); "STATUS                 :   "; RSGAJI!Status
Printer.Print Tab(5); "JUMLAH ANAK            :  "; RSGAJI!JMLANAK
Printer.Print Tab(5); "JUMLAH HARI KERJA      :  "; RSGAJI!masuk
Printer.Print Tab(5); "JUMLAH JAM LEMBUR      :  "; RSGAJI!LEMBUR ' dibuang kata JAM
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "GAJI POKOK             :   "; RKanan(RSGAJI!GAPOK, "##,###,###")
Printer.Print Tab(5); "TUNJANGAN JABATAN      :   "; RKanan(RSGAJI!TJJABATAN, "##,###,###")
Printer.Print Tab(5); "TUNJANGAN SUAMI / ISTRI:   ";

If RSGAJI!TJKELUARGA = 0 Then
    Printer.Print Tab(40); RSGAJI!TJKELUARGA
Else
    Printer.Print Tab(32); RKanan(RSGAJI!TJKELUARGA, "##,###,###")
End If

Printer.Print Tab(5); "TUNJANGAN ANAK         :   ";
If RSGAJI!TJANAK = 0 Then
    Printer.Print Tab(40); RSGAJI!TJANAK
Else
    Printer.Print Tab(32); RKanan(RSGAJI!TJANAK, "##,###,###")
End If

Printer.Print Tab(5); "UANG MAKAN             :   "; RKanan(RSGAJI!UANGMAKAN, "##,###,###")

Printer.Print Tab(5); "UANG LEMBUR            :   ";
If RSGAJI!UANGLEMBUR = 0 Then
    Printer.Print Tab(40); RSGAJI!UANGLEMBUR
Else
    Printer.Print Tab(32); RKanan(RSGAJI!UANGLEMBUR, "##,###,###");
End If

Printer.Print Tab(5); "ASKES                  :   "; RKanan(RSGAJI!ASKES, "##,###,###")
Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "TOTAL PENDAPATAN       :   "; RKanan(RSGAJI!PENDAPATAN, "##,###,###")
Printer.Print Tab(5); MGrs

Printer.Print Tab(5); "POTONGAN               :   ";
If RSGAJI!POTONGAN = 0 Then
    Printer.Print Tab(40); RSGAJI!POTONGAN
Else
    Printer.Print Tab(32); RKanan(RSGAJI!POTONGAN, "##,###,###")
End If

Printer.Print Tab(5); MGrs
Printer.Print Tab(5); "TOTAL GAJI             :   "; RKanan(RSGAJI!TotalGAJI, "##,###,###")
Printer.Print Tab(5); MGrs
Printer.Print
Printer.Print
Conn.Close
Printer.EndDoc
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function


