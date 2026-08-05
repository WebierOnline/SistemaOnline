Attribute VB_Name = "mCampoCelular"
Function CampoCelular(obj As Object, Keyasc As Integer)

If Not ((Keyasc >= Asc("0") And Keyasc <= Asc("9")) Or Keyasc = 8) Then
   Keyasc = 0
   Exit Function
End If

If Keyasc <> 13 Then
  
   If Len(obj.Text) = 2 Then
      obj.Text = obj.Text + ")"
      obj.SelStart = Len(obj.Text)
   End If
   
   If Len(obj.Text) = 3 Then
      obj.Text = obj.Text + " "
      obj.SelStart = Len(obj.Text)
   End If
   
   If Len(obj.Text) = 9 Then
      obj.Text = obj.Text + "-"
      obj.SelStart = Len(obj.Text)
   End If
      
End If

End Function

