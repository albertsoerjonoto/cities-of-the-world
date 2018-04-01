VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Kota-kota di Indonesia"
   ClientHeight    =   7320
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Hapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2520
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Tambahkan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   14
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      TabIndex        =   13
      Top             =   4200
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "DEL >"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   3
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< DEL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Kota-Kota.frx":0000
      Left            =   5400
      List            =   "Kota-Kota.frx":0002
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Kota-Kota.frx":0004
      Left            =   240
      List            =   "Kota-Kota.frx":0006
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Gambar Kota"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   3960
      Top             =   4200
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "INPUT DATA :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "JUMLAH DATA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "JUMLAH DATA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListIndex = -1 Then
    x = MsgBox("Tidak Ada kota yang dipilih...!", vbExclamation, "Warning")
Exit Sub
End If
List1.RemoveItem (List1.ListIndex)
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command2_Click()
If List1.ListIndex = -1 Then
    x = MsgBox("Tidak Ada kota yang dipilih...!", vbExclamation, "Warning")
Exit Sub
End If
List2.AddItem (List1.Text)
List1.RemoveItem (List1.ListIndex)
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command3_Click()
For z = 0 To List1.ListCount - 1
List2.AddItem List1.List(z)
Next
List1.Clear
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command4_Click()
For y = 0 To List2.ListCount - 1
List1.AddItem List2.List(y)
Next
List2.Clear
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command5_Click()
If List2.ListIndex = -1 Then
    x = MsgBox("Tidak Ada kota yang dipilih...!", vbExclamation, "Warning")
Exit Sub
End If
List1.AddItem (List2.Text)
List2.RemoveItem (List2.ListIndex)
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command6_Click()
If List2.ListIndex = -1 Then
    x = MsgBox("Tidak Ada kota yang dipilih...!", vbExclamation, "Warning")
Exit Sub
End If
List2.RemoveItem (List2.ListIndex)
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Form_Load()
List1.AddItem "Surabaya"
List1.AddItem "Jakarta"
List1.AddItem "Bandung"
List1.AddItem "Tegal"
List1.AddItem "Tangerang"
List1.AddItem "Bekasi"
List1.AddItem "Depok"
List1.AddItem "Bogor"
List1.AddItem "Pemalang"
List1.AddItem "Denpasar"
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub

Private Sub Hapus_Click()
Text1.Text = ""
End Sub

Private Sub List1_Click()
If List1.Text = "Surabaya" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Surabaya.jpg")
ElseIf List1.Text = "Pemalang" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Pemalang.jpg")
ElseIf List1.Text = "Jakarta" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Jakarta.jpg")
ElseIf List1.Text = "Denpasar" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Denpasar.jpg")
ElseIf List1.Text = "Bekasi" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bekasi.jpg")
ElseIf List1.Text = "Tegal" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Tegal.jpg")
ElseIf List1.Text = "Tangerang" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Tangerang.jpg")
ElseIf List1.Text = "Bogor" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bogor.jpg")
ElseIf List1.Text = "Depok" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Depok.jpg")
ElseIf List1.Text = "Bandung" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bandung.jpg")
Else
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\0.jpg")
End If
End Sub

Private Sub List2_Click()
If List2.Text = "Surabaya" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Surabaya.jpg")
ElseIf List2.Text = "Pemalang" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Pemalang.jpg")
ElseIf List2.Text = "Jakarta" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Jakarta.jpg")
ElseIf List2.Text = "Denpasar" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Denpasar.jpg")
ElseIf List2.Text = "Bekasi" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bekasi.jpg")
ElseIf List2.Text = "Tegal" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Tegal.jpg")
ElseIf List2.Text = "Tangerang" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Tangerang.jpg")
ElseIf List2.Text = "Bogor" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bogor.jpg")
ElseIf List2.Text = "Depok" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Depok.jpg")
ElseIf List2.Text = "Bandung" Then
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\Bandung.jpg")
Else
    Image1.Picture = LoadPicture("C:\Documents and Settings\User\Desktop\Tugas Komputer\VB\ListBox\Foto kota-kota\0.jpg")
End If
End Sub


Private Sub Command7_Click()
If Text1 = "" Then
    x = MsgBox("Tuliskan nama Kota!", vbExclamation, "Warning")
    Add.SetFocus
Else
    List1.AddItem Text1.Text
    List1.ListIndex = List1.ListCount - 1

End If
Label3.Caption = List1.ListCount
Label5.Caption = List2.ListCount
End Sub
    
    
