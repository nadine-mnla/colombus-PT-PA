VERSION 5.00
Begin VB.Form frmGuest
   Caption         =   "Colombus - Guest Entry"
   ClientHeight    =   2000
   ClientLeft      =   100
   ClientTop       =   100
   ClientWidth     =   4000
   Begin VB.TextBox txtNama
      Height          =   300
      Left            =   200
      Top             =   200
      Width           =   2000
   End
   Begin VB.TextBox txtKTP
      Height          =   300
      Left            =   200
      Top             =   600
      Width           =   2000
   End
   Begin VB.TextBox txtKamarID
      Height          =   300
      Left            =   200
      Top             =   1000
      Width           =   2000
   End
   Begin VB.CommandButton cmdSave
      Caption         =   "Simpan"
      Height          =   350
      Left            =   200
      Top             =   1400
      Width           =   1500
   End
End
Attribute VB_Name = "frmGuest"

Private Sub cmdSave_Click()
    Call OpenConnection
    Dim sql As String
    sql = "INSERT INTO Tamu (Nama, NoKTP, KamarID) VALUES ('" & txtNama.Text & "', '" & txtKTP.Text & "', " & txtKamarID.Text & ")"
    conn.Execute sql
    MsgBox "Tamu berhasil disimpan."
    Call CloseConnection
End Sub