VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Terbilang Indonesia"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function TerbilangBulat(strAngka As String, _
Optional MataUang As String = "rupiah") As String
   Dim strJmlHuruf$, intPecahan As Integer
   Dim strPecahan$, Urai$, Bil1$, strTot$, Bil2$
   Dim X As Integer, Y As Integer, z As Integer
   On Error GoTo Pesan
   Dim strValid As String, huruf As String * 1
   Dim i As Integer
   'Periksa setiap karakter yg diketikkan ke kotak
   'UserID
   strValid = "1234567890"
   For i% = 1 To Len(strAngka)
     huruf = Chr(Asc(Mid(strAngka, i%, 1)))
     If InStr(strValid, huruf) = 0 Then
       Set AngkaTerbilang = Nothing
       MsgBox "Harus karakter angka!", _
              vbCritical, "Karakter Tidak Valid"
       Exit Function
     End If
   Next i%
    
   If strAngka = "" Then Exit Function
   If Len(Trim(strAngka)) > 15 Then GoTo Pesan
   strJmlHuruf = LTrim(strAngka)
   'intPecahan = Val(Right(Mid(strAngka, 15, 2), 2))
   
   If (intPecahan = 0) Then
      strPecahan = ""
   Else
      'strPecahan = LTrim(Str(intPecahan)) + "/100 "
      strPecahan = ""
   End If

   X = 0
   Y = 0
   Urai = ""
   While (X < Len(strJmlHuruf))
     X = X + 1
     strTot = Mid(strJmlHuruf, X, 1)
     Y = Y + Val(strTot)
     z = Len(strJmlHuruf) - X + 1
     Select Case Val(strTot)
     Case 1
       If (z = 1 Or z = 7 Or z = 10 Or z = 13) Then
          Bil1 = "satu "
       ElseIf (z = 4) Then
          If (X = 1) Then
             Bil1 = "se"
          Else
             Bil1 = "satu "
          End If
       ElseIf (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
          X = X + 1
          strTot = Mid(strJmlHuruf, X, 1)
          z = Len(strJmlHuruf) - X + 1
          Bil2 = ""
        
          Select Case Val(strTot)
                 Case 0:   Bil1 = "sepuluh "
                 Case 1:   Bil1 = "sebelas "
                 Case 2:   Bil1 = "dua belas "
                 Case 3:   Bil1 = "tiga belas "
                 Case 4:   Bil1 = "empat belas "
                 Case 5:   Bil1 = "lima belas "
                 Case 6:   Bil1 = "enam belas "
                 Case 7:   Bil1 = "tujuh belas "
                 Case 8:   Bil1 = "delapan belas "
                 Case 9:   Bil1 = "sembilan belas "
          End Select
       Else
          Bil1 = "se"
       End If
     Case 2:   Bil1 = "dua "
     Case 3:   Bil1 = "tiga "
     Case 4:   Bil1 = "empat "
     Case 5:   Bil1 = "lima "
     Case 6:   Bil1 = "enam "
     Case 7:   Bil1 = "tujuh "
     Case 8:   Bil1 = "delapan "
     Case 9:   Bil1 = "sembilan "
     Case Else
               Bil1 = ""
     End Select

     If (Val(strTot) > 0) Then
        If (z = 2 Or z = 5 Or z = 8 Or z = 11 Or z = 14) Then
           Bil2 = "puluh "
        ElseIf (z = 3 Or z = 6 Or z = 9 Or z = 12 Or z = 15) Then
           Bil2 = "ratus "
        Else
           Bil2 = ""
        End If
     Else
        Bil2 = ""
     End If
    
     If (Y > 0) Then
        Select Case z
               Case 4:    Bil2 = Bil2 + "ribu "
                          Y = 0
               Case 7:    Bil2 = Bil2 + "juta "
                          Y = 0
               Case 10:   Bil2 = Bil2 + "milyar "
                          Y = 0
               Case 13:   Bil2 = Bil2 + "trilyun "
                          Y = 0
        End Select
     End If
     Urai = Urai + Bil1 + Bil2
   Wend
   Urai = Urai + strPecahan
   TerbilangBulat = (Urai & MataUang)
   Exit Function
Pesan:
   TerbilangBulat = "(maksimal 15 digit)"
End Function

Private Sub Text1_Change()
   Text2.Text = TerbilangBulat(Text1.Text)
End Sub


