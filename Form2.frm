VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H000080FF&
   Caption         =   "apip kurniawan"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   ForeColor       =   &H0000C000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   5565
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak Nota"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4920
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid grdtabel 
      Height          =   2295
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtptgl 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   6815745
      CurrentDate     =   42811
   End
   Begin VB.TextBox txtkasir 
      Height          =   285
      Left            =   6840
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtnota 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Kasir :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "No Nota :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Aplikasi Penjualan Tunai Toko Orang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.Recordset
Dim vtotal As Single

Private Sub cmdcetak_Click()
Open "d:\nota.txt" For Output As #1
Print #1, "                INDOMART"
Print #1, "       Jl. Raya Setu Kp.Bahagia Tambun"
Print #1, ""
Print #1, "    No : " + txtnota.Text + " Tanggal : " + Str(dtptgl.Value) + " Kasir : " + txtkasir.Text
Print #1, "---------------------------------------------"
Print #1, "No"; Tab(4); "Nama Barang"; Tab(26); "Harga"; Tab(35); "Qty"; Tab(40); "Jumlah"
Print #1, "============================================="
RS.MoveFirst
For X = 1 To RS.RecordCount
  If RS.Fields("nama_barang") <> "" Then
  no = no + 1
  'mencetak nilai variabel ke dalam file laporan
  Print #1, no; Tab(4); RS.Fields(0); Tab(26);
  Print #1, RS.Fields(1); Tab(35);
  Print #1, RS.Fields(2); Tab(40);
  Print #1, RS.Fields(3)
  End If


  RS.MoveNext
  
Next X
Print #1, "----------------------------------------------"
Print #1, Tab(26); "Total : "; Tab(35); Format(vtotal, "currency")
Print #1, ""
Print #1, ""
Print #1, Tab(10); "terimakasih , semoga anda puas"
Close #1

nota.Show
End Sub

Private Sub cmdhapus_Click()
lbltotal.Caption = ""
RS.MoveFirst
For X = 1 To RS.RecordCount
  RS.Delete
  RS.Update
  RS.MoveNext
Next X
 
For X = 1 To 20
  RS.AddNew
Next X

RS.MoveFirst
grdtabel.SetFocus

End Sub

Private Sub Form_Load()
'memberikan nilai variabel dengan objek data
Set RS = New ADODB.Recordset

'membuat struktur tabel
RS.Fields.Append "Nama_Barang", adVarChar, 50
RS.Fields.Append "Harga", adSingle, 50
RS.Fields.Append "Qty", adSingle, 50
RS.Fields.Append "Jumlah", adSingle, 50

'mengaktifkan variabel data objek
RS.Open

'memasukan objek data ke dalam objek grid
Set grdtabel.DataSource = RS

'mengatur tampilan tabel
'a. mengatur lebar kolom
grdtabel.Columns(0).Width = 3000

'memberi baris data kosong
For X = 1 To 20
    RS.AddNew
    Next X
RS.MoveFirst

End Sub

Private Sub grdtabel_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
Case 0:
  grdtabel.Col = 1
Case 1:
  grdtabel.Col = 2
Case 2:
  RS.Fields("Jumlah") = RS.Fields("Harga") * RS.Fields("Qty")
  vtotal = vtotal + (RS.Fields("Harga") * RS.Fields("Qty"))
  lbltotal.Caption = Format(vtotal, "currency")
  grdtabel.Col = 0
  RS.MoveNext
End Select
  
End Sub

