VERSION 5.00
Begin VB.Form FrmTerbilang 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terbilang"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmTerbilang.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtAngka 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Terbilang :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Masukkan angka :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label LblHuruf 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   6720
      WordWrap        =   -1  'True
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "FrmTerbilang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Help_Click()
MsgBox ("Masukkan angka ke dalam textbox maka program akan menampilkan terbilangnya." & vbNewLine & _
    "Angka yang dimasukkan harus berupa bilangan bulat positif." & vbNewLine & _
    "Angka yang dimasukkan tidak boleh lebih dari 9 digit."), vbInformation, "Panduan"
    
End Sub
Private Sub TxtAngka_Change()

'deklarasi variabel
Dim A, i, J As Integer
Dim Digit(9), Bagian(3) As Integer
Dim Huruf(9), Terbilang(3), HASIL As String

'membaca angka
A = Val(TxtAngka.Text)

'mencegah error
If A < 0 Or A > 999999999 Then
    MsgBox ("Input harus berupa bilangan positif maksimal 9 digit."), vbExclamation
    Exit Sub
End If

'membagi menjadi 3 3
For i = 1 To 3
    Bagian(i) = Right(A, 3)
    A = (A - Bagian(i)) / 1000
Next i

For i = 1 To 3
    'memisahkan per digit
        'satuan
        Digit(1) = Right(Bagian(i), 1)
        'puluhan
        If Bagian(i) > 9 Then
            Digit(2) = Right(Bagian(i), 2)
            Digit(2) = Left(Digit(2), 1)
            Else
            Digit(2) = 0
        End If
        'ratusan
        If Bagian(i) > 99 Then
            Digit(3) = Left(Bagian(i), 1)
            Else
            Digit(3) = 0
        End If

    'menulis digit dalam huruf
    For N = 1 To 3
        Select Case Digit(N)
            Case 1
            Huruf(N) = "Satu"
            Case 2
            Huruf(N) = "Dua"
            Case 3
            Huruf(N) = "Tiga"
            Case 4
            Huruf(N) = "Empat"
            Case 5
            Huruf(N) = "Lima"
            Case 6
            Huruf(N) = "Enam"
            Case 7
            Huruf(N) = "Tujuh"
            Case 8
            Huruf(N) = "Delapan"
            Case 9
            Huruf(N) = "Sembilan"
            Case 0
            Huruf(N) = Empty
        End Select
    Next N

    'menambahkan puluh dan ratus
        If Digit(2) > 0 Then
            Huruf(2) = Huruf(2) + " Puluh "
            Else
            Huruf(2) = ""
        End If
    
        If Digit(3) > 0 Then
            Huruf(3) = Huruf(3) + " Ratus "
            Else
            Huruf(3) = ""
        End If

    'mengatur angka satu
        'mengatur ratusan
        If Digit(3) = 1 Then
            Huruf(3) = "Seratus "
        End If
    
        'mengatur belasan
        If Digit(2) = 1 And Digit(1) > 1 Then
            Huruf(2) = Huruf(1) + " "
            Huruf(1) = "Belas"
        End If
    
        'mengatur angka sebelas
        If Digit(2) = 1 And Digit(1) = 1 Then
            Huruf(2) = "Se"
            Huruf(1) = "belas"
        End If
        
        'mengatur angka sepuluh
        If Digit(2) = 1 And Digit(1) = 0 Then
            Huruf(2) = "Se"
            Huruf(1) = "puluh"
        End If
        
    'menggabungkan kalimat
    Terbilang(i) = Huruf(1)

    For N = 2 To 3
        Terbilang(i) = Huruf(N) + Terbilang(i)
    Next N
Next i

'menambahkan juta
Terbilang(1) = Terbilang(1)
Terbilang(2) = Terbilang(2) + " Ribu "
Terbilang(3) = Terbilang(3) + " Juta "

'mengatur angka nol 000
If Terbilang(3) = " Juta " Then
    Terbilang(3) = ""
End If
If Terbilang(2) = " Ribu " Then
    Terbilang(2) = ""
End If
If Terbilang(2) = "Satu Ribu " Then
    Terbilang(2) = "Seribu "
End If

'menggabungkan terbilang
HASIL = Terbilang(1)
For i = 2 To 3
    HASIL = Terbilang(i) + HASIL
Next i

'menampilkan hasil
LblHuruf.Caption = HASIL
End Sub
