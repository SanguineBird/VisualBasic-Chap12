Public Class Form1
    Private rect As New Rectangle
    Private carpet As New Carpet
    Private decTotal As Decimal

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim strPricePer As String
        Dim strWidth As String
        Dim strLength As String

        txtArea.Text = ""
        txtCost.Text = ""

        'collect user input
        carpet.Color = txtColor.Text
        carpet.Style = txtStyle.Text
        strPricePer = txtPricePer.Text
        strWidth = txtWidth.Text
        strLength = txtLength.Text

        'validate user input and send to outer classes
        If IsValid(strPricePer) Then
            carpet.PricePer = CDec(strPricePer)
        Else
            carpet.PricePer = 0
        End If

        If IsValid(strWidth) Then
            rect.Width = CSng(strWidth)
        Else
            rect.Width = 0
        End If

        If IsValid(strLength) Then
            rect.Length = CSng(strLength)
        Else
            rect.Length = 0
        End If

        'calculate and display area
        rect.CalcArea()
        txtArea.Text = rect.Area.ToString()

        'calculate and display total cost
        decTotal = CDec(rect.Area * carpet.PricePer)
        txtCost.Text = decTotal.ToString("C")
    End Sub

    Private Function IsValid(strInput As String) As Boolean
        Dim blnResult As Boolean

        If Not String.IsNullOrEmpty(strInput) Then
            If IsNumeric(strInput) Then
                If CDec(strInput) >= 0 Then
                    blnResult = True
                Else
                    MessageBox.Show("Error: Price, Length, and Width input must be a positive number.")
                    blnResult = False
                End If
            Else
                MessageBox.Show("Error: Price, Length, and Width input must be numeric.")
                blnResult = False
            End If
        Else
            MessageBox.Show("Error: Price, Length, and Width input cannot be blank.")
            blnResult = False
        End If

        Return blnResult
    End Function

    Private Sub clearAll()
        carpet.Color = ""
        carpet.Style = ""
        carpet.PricePer = 0
        rect.Width = 0
        rect.Length = 0
        decTotal = 0

        txtColor.Text = ""
        txtStyle.Text = ""
        txtPricePer.Text = ""
        txtWidth.Text = ""
        txtLength.Text = ""
        txtArea.Text = ""
        txtCost.Text = ""
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearAll()
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub
End Class

Public Class Rectangle
    Private sngWidth As Single
    Private sngLength As Single
    Private sngArea As Single

    Public Property Width As Single
        Get
            Return sngWidth
        End Get
        Set(value As Single)
            sngWidth = value
        End Set
    End Property
    Public Property Length As Single
        Get
            Return sngLength
        End Get
        Set(value As Single)
            sngLength = value
        End Set
    End Property
    Public ReadOnly Property Area As Single
        Get
            Return sngArea
        End Get
    End Property

    Public Sub CalcArea()
        sngArea = sngWidth * sngLength
    End Sub
End Class

Public Class Carpet
    Public Property Color As String
    Public Property Style As String
    Public Property PricePer As Decimal
End Class
