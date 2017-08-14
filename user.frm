VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form user 
   BackColor       =   &H00FFFF80&
   Caption         =   "user"
   ClientHeight    =   8565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14520
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton baru 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   5400
      Width           =   1815
   End
   Begin VB.ComboBox hakakses 
      Height          =   315
      Left            =   3000
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6255
      Left            =   6000
      TabIndex        =   13
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11033
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.CommandButton keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton edit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton batal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton simpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox password 
      Height          =   405
      Left            =   3000
      TabIndex        =   7
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox kodeuser 
      Height          =   405
      Left            =   3000
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox namauser 
      Height          =   405
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF80&
      Caption         =   "password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF80&
      Caption         =   "hak akses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "kode user"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      Caption         =   "nama user"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      Caption         =   "USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktif()
    kodeuser.Enabled = True
    namauser.Enabled = True
    password.Enabled = True
    hakakses.Enabled = True
    baru.Enabled = True
    
End Sub

Private Sub tidakaktif()
    kodeuser.Enabled = True
    namauser.Enabled = True
    password.Enabled = True
    hakakses.Enabled = True
End Sub

Private Sub kosong()
    kodeuser.Text = ""
    namauser.Text = ""
    password.Text = ""
    hakakses.Text = ""
End Sub

Private Sub tampil()
     kodeuser.Text = rsuser!kode_user
    namauser.Text = rsuser!nama_user
    password.Text = rsuser!password
    hakakses.Text = rsuser!hak_akses
End Sub



Private Sub cancel_Click()
 Call Form_Activate
 Save.Enabled = True
End Sub

Private Sub batal_Click()
Call Form_Activate
 baru.Enabled = True
End Sub

Private Sub DataGrid1_Click()

    Call tampil
    Call aktif
    namauser.Enabled = True
    namauser.Enabled = False
    simpan.Enabled = False
    baru.Enabled = False
    edit.SetFocus
End Sub

Private Sub hakakses_Click()
nomer = "select  Kode_user from user where left (Kode_user, 3)='" & "KSR" & "'order by Kode_user desc"
Set rsuser = conn.Execute(nomer, , adCmdText)
Dim urutan As String * 10
If Not rsuser.EOF Then
   hitung = Right(rsuser!kode_user, 2) + 1
 Select Case hakakses
 Case "KASIR"
 kodeuser.Text = "KSR" + "0" & (Trim(Str(hitung)))
 namauser.Enabled = True
namauser.SetFocus
Exit Sub
 End Select
 End If
nomer = "select  Kode_user from user where left (Kode_user, 3)='" & "ADM" & "'order by Kode_user desc"
Set rsuser = conn.Execute(nomer, , adCmdText)
If Not rsuser.EOF Then
   hitung = Right(rsuser!kode_user, 2) + 1
 Select Case hakakses
 Case "ADMIN"
 kodeuser.Text = "ADM" + "0" & (Trim(Str(hitung)))
 namauser.Enabled = True
namauser.SetFocus
Exit Sub
 End Select
 End If
nomer = "select  Kode_user from user where left (Kode_user, 3)='" & "KRU" & "'order by Kode_user desc"
Set rsuser = conn.Execute(nomer, , adCmdText)

If Not rsuser.EOF Then
   hitung = Right(rsuser!kode_user, 2) + 1
 Select Case hakakses
 Case "KURIR"
 kodeuser.Text = "KRU" + "0" & (Trim(Str(hitung)))
 namauser.Enabled = True
namauser.SetFocus
 End Select
 End If
End Sub



Private Sub Combo1_Change()

End Sub

Private Sub hapus_Click()
Dim hapus As String
    hapus = "DELETE FROM user WHERE nama_user = '" & namauser.Text & "'"
    conn.Execute hapus
    MsgBox "Data Berhasil Dihapus !", vbOKOnly, "Info"
    Call Form_Activate
    simpan.Enabled = True
    baru.Enabled = True
    
End Sub

Private Sub edit_Click()
Dim update As String
    update = "UPDATE user SET password= '" & password.Text & "',hak_akses= '" & hakakses.Text & "',kode_user= '" & kodeuser.Text & "' WHERE nama_user = '" & namauser.Text & "'"
    conn.Execute update
    MsgBox "Data berhasil diubah !", vbOKOnly, "Info"
    Call Form_Activate
    simpan.Enabled = True
    baru.Enabled = True
End Sub

Private Sub exit_click()
Unload Me
End Sub

Private Sub Form_Activate()
    Call kosong
    Call aktif
    Call koneksi
    baru.Enabled = True
    rsuser.Open "SELECT * FROM user", conn
    Set DataGrid1.DataSource = rsuser.DataSource
    
    baru.SetFocus
    
DataGrid1.Columns(2).Width = 3000
End Sub

Private Sub baru_Click()
Call aktif
hakakses.Enabled = True
Call kosong
baru.Enabled = True
kodeuser.SetFocus
End Sub

Private Sub Form_Load()
hakakses.AddItem "ADMIN"
hakakses.AddItem "KASIR"
End Sub

Private Sub grid_user_Click()
    Call tampil
    Call aktif
    kodeuser.Enabled = False
    namauser.Enabled = False
    Save.Enabled = False
    baru.Enabled = False
    edit.SetFocus
End Sub

Private Sub keluar_Click()
Unload Me
End Sub

Private Sub namauser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
password.SetFocus

If Not (KeyAscii >= Asc("a") & Chr(13) And KeyAscii <= Asc("9") & Chr(13) Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace) Then
KeyAscii = a
End If
End If
End Sub


    


Private Sub simpan_Click()
Dim create As String
    create = "INSERT INTO user (nama_user,hak_akses,password,kode_user) values ('" & namauser.Text & "','" & hakakses.Text & "','" & password.Text & "','" & kodeuser.Text & "')"
    conn.Execute create
    MsgBox "Data berhasil disimpan !", vbOKOnly, "Info"
    Call Form_Activate
End Sub
