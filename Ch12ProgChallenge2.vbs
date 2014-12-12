'Ch 12 Prog Challenge 2
'Carpet Price Calculator

Private Class Rectangle
  Public Property SngWidth as Single
  Public Property SngLength as Single
  ReadOnly Property SngArea as Single
  
  Private Sub CalcArea
    SngArea = SngWidth x SngLength
  End Sub
End Class

Private Class Carpet
  Public Property StrColor as String
  Public Property StrStyle as String
  Public Property DecPrice as Decimal
  Dim rect as New Rectangle
  Dim curTotal as Currency
  
  Sub handles btnCalculate
    clearAll()
    
    'collect user input
    StrColor = txtColor.Text
    StrStyle = txtStyle.Text
    If isValid(txtPrice.Text) Then
      DecPricePer = CDec(txtPrice.Text)
    Else
      DecPricePer = 0
    End If
    
    If isValid(txtWidth.Text) Then
      rect.SngWidth = CSng(txtWidth.Text)
    Else
      rect.SngWidth = 0
    End If
    
    If isValid(txtLength.Text) Then
      rect.SngLength = CSng(txtLength.Text)
    Else
      rect.SngLength = 0
    End If
    
    'calculate and display area
    rect.CalcArea()
    txtArea.Text = rect.SngArea.toString()
    
    'calculate and display total cost
    curTotal = rect.SngArea * DecPricePer
    txtCost.Text = curTotal.toString("C")
  End Sub
  
  Sub handles btnClear
    clearAll()
  End Sub
  
  Sub handles btnExit
    Me.Close()
  End Sub
  
  Private Sub clearAll
    StrColor = ""
    StrStyle = ""
    DecPricePer = 0
    rect.SngWidth = 0
    rect.SngLength = 0
    curTotal = 0
    
    txtColor.Text = ""
    txtStyle.Text = ""
    txtPricePer.Text = ""
    txtWidth.Text = ""
    txtLength.Text = ""
    txtArea.Text = ""
    txtCost.Text = ""
  End Sub
  
  Function IsValid (strInput as String) as Boolean
    Dim blnResult as Boolean = 0
    
    If ((strInput <> "") && (isNumeric.strInput())) Then
      If (CDec(strInput) >= 0) Then
        blnResult = 1
      Else
        MessageBox.Show("Error: Invalid Input.")
        blnResult = 0
      End If
    Else
      MessageBox.Show("Error: Invalid Input.")
      blnResult = 0
    End If
    
    Return blnResult
  End Function
End Class
