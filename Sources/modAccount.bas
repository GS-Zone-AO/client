Attribute VB_Name = "modAccount"
Option Explicit

Public Type AccountUser
    Name As String
    Body As Integer
    Head As Integer
    Weapon As Integer
    Shield As Integer
    Helmet As Integer
    Class As Byte
    Race As Byte
    Map As String
    Level As Byte
    Gold As Long
    LastConnect As String
    Criminal As Boolean
    Dead As Boolean
    GameMaster As Boolean
End Type

Public AccountName As String
Public NumberOfCharacters As Byte
Public Characters() As AccountUser

Public Const MaxCharPerAccount As Byte = 20

Sub LoadCharacters()

    Call frmConnect.lPersonajes.Clear
    Dim i As Byte
    For i = 1 To NumberOfCharacters
        Call frmConnect.lPersonajes.AddItem(Characters(i).Name & " [" & Characters(i).Level & "]")
    Next
    frmConnect.lPersonajes.Refresh
    frmConnect.cCrearPJ.Enabled = (NumberOfCharacters < MaxCharPerAccount)
    
End Sub
