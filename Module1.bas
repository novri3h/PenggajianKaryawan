Attribute VB_Name = "Module1"


Public Conn As New ADODB.Connection
Public RSGAJI As ADODB.Recordset
Public RSKASIR As ADODB.Recordset
Public RSGOL As ADODB.Recordset
Public RSJabatan As ADODB.Recordset
Public RSPegawai As ADODB.Recordset
Public RSMASTER As ADODB.Recordset
Public RSTEMPORER As ADODB.Recordset
Public RSABSEN As ADODB.Recordset

Public Sub BukaDB()
Set Conn = New ADODB.Connection
Set RSGAJI = New ADODB.Recordset
Set RSKASIR = New ADODB.Recordset
Set RSGOL = New ADODB.Recordset
Set RSJabatan = New ADODB.Recordset
Set RSPegawai = New ADODB.Recordset
Set RSMASTER = New ADODB.Recordset
Set RSTEMPORER = New ADODB.Recordset
Set RSABSEN = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBGaji.mdb"
End Sub


