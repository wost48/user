VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox DataGrid1 
      Height          =   3135
      Left            =   1080
      ScaleHeight     =   3075
      ScaleWidth      =   12915
      TabIndex        =   20
      Top             =   5280
      Width           =   12975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   12480
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cari"
      Height          =   375
      Left            =   10680
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Penghitung Barang"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   11880
      TabIndex        =   12
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11880
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11880
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Harga Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "PROGRAM PENJUALAN BUSANA MUSLIM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim caridata As String

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
caridata = InputBox("Masukkan Nama Barang", "Cari Data")
If rs.State = adStateOpen Then rs.Close
rs.Open
"select * from barang where Nama_Barang="&caridata&"",con,adOpenDynamic,
adLockOptimistic
Set DataGrid1.DataSource = rs
End Sub

Private Sub bukabarang()
    If rs.State = adStateOpen Then rs.Close
    rs.Open "select * from Barang ", con, adOpenDynamic, adLockOptimistic
    Set DataGrid1.DataSource = rs
End Sub

Private Sub tampilbarang()
  With rs
    Text1.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Kode_Barang), "", !Kode_Barang))
    Text2.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Nama_Barang), "", !Nama_Barang))
    Text3.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Jenis_Barang), "", !Jenis_Barang))
    Text4.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Harga_Barang), "", !Harga_Barang))
    Text5.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Jumlah_Barang), "", !Jumlah_Barang))
    Text6.Text = IIf(.BOF Or .EOF, "", IIf(IsNull(!Keterangan), "", !Keterangan))
  End With
End Sub

Private Sub Command7_Click()
bukabarang
tampilbarang
End Sub

Private Sub Form_Load()
'panggil procedure konek
konekdb
'seleksi tabel
penjualanbarang
'konekkan dengan object rs
    bukabarang
    tampilbarang
Set DataGrid1.DataSource = rs
End Sub

'tambah data
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = "0"
Text5.Text = "0"
Text6.Text = ""
Text1.SetFocus
End Sub

'simpan data
Private Sub Command2_Click()
Call insert(Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, Text6.Text)
End Sub

'hapus data
Private Sub Command3_Click()
If Not (rs.EOF Or rs.BOF) Then
rs.Delete
Else
MsgBox "data tidak ada"
End If
End Sub
