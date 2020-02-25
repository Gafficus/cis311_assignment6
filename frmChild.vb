#Disable Warning IDE1006 ' Naming Styles
Public Class frmChild
    Private txtSelectedTextBox As TextBox
    Private Sub btnPutNumberIntoText(sender As Button, e As EventArgs) Handles btn0.Click, btn1.Click, btn3.Click,
                                                                               btn4.Click, btn5.Click, btn6.Click,
                                                                               btn7.Click, btn8.Click, btn9.Click,
                                                                               btnDot.Click, btn2.Click
        Try
            txtSelectedTextBox.Text &= sender.Tag.ToString()
        Catch ex As Exception
            Debug.WriteLine("Exception caught on inputting a number.")
        End Try

    End Sub

    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click

    End Sub

    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        Dim txtBoxes = From txtBox In Me.Controls
                       Where TypeOf (txtBox) Is TextBox
                       Select txtBox
        For Each txtBox As TextBox In txtBoxes
            txtBox.Clear()
            txtBox.Text = 0
        Next

    End Sub

    Private Sub btnCE_Click(sender As Object, e As EventArgs) Handles btnCE.Click
        If Not IsNothing(txtSelectedTextBox) Then
            txtSelectedTextBox.Clear()
            txtSelectedTextBox.Text = "0"
        Else
            MessageBox.Show("Please select a text box.")
        End If

    End Sub

    Private Sub txtBox_GotFocus(sender As TextBox, e As EventArgs) Handles txtFirstVariable.GotFocus,
                                                                           txtSecndVariable.GotFocus,
                                                                           txtThirdVariable.GotFocus

        txtSelectedTextBox = sender
    End Sub

    Private Sub frmChild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim lstUdtShapes As New List(Of clsShape)
        lstUdtShapes.Add(New clsRectangle)
        lstUdtShapes.Add(New clsSquare)
        lstUdtShapes.Add(New clsRightTriangle)
        lstUdtShapes.Add(New clsCircle)
        lstUdtShapes.Add(New clsCube)
        lstUdtShapes.Add(New clsSphere)
        lstUdtShapes.Add(New clsCylinder)
        lstUdtShapes.Add(New clsCone)

        For Each shape In lstUdtShapes
            lstShapes.Items.Add(shape)
        Next
    End Sub

    Private Sub lstFormulas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstFormulas.SelectedIndexChanged

    End Sub

    Private Sub lstShapes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstShapes.SelectedIndexChanged
        lstFormulas.Items.Clear()
        With lstShapes.SelectedItem
            For Each s In .GetstrFormulaTypes()
                lstFormulas.Items.Add(s)
            Next
        End With
    End Sub

#Enable Warning IDE1006 ' Naming Styles

End Class