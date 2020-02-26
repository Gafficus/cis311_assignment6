#Disable Warning IDE1006 ' Naming Styles
Imports System.ComponentModel
'------------------------------------------------------------
'-                File Name : frmChild.vb                     - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 20 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the event handlers for the child form -
'- of the mdi parent.
'------------------------------------------------------------	
Public Class frmChild
    Enum formula
        PerimeterCicumferenceVolume = 0
        AreaSurfaceArea = 1
    End Enum
    Private txtSelectedTextBox As TextBox
    '------------------------------------------------------------
    '-                Subprogram Name: btnPutNumberIntoText     -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when pressing the number       -
    '- calculator buttons. It will place the corresponding      -
    '- number into the last text box selected.                  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub btnPutNumberIntoText(sender As Button, e As EventArgs) Handles btn0.Click, btn1.Click, btn3.Click,
                                                                               btn4.Click, btn5.Click, btn6.Click,
                                                                               btn7.Click, btn8.Click, btn9.Click,
                                                                               btnDot.Click, btn2.Click
        'Prevents exceptions from being thrown when selecting a number
        Try
            txtSelectedTextBox.Text &= sender.Tag.ToString()
        Catch ex As Exception
            Debug.WriteLine("Exception caught on inputting a number.")
        End Try

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnCalculate_Click     -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when pressing the calculate    -
    '- button. It will take the numbers in the text boxes and   -
    '- then send the variables to the shapes formula, depending -
    '- On which formula Is selected
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):              -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblAnswer - holds the answer from the formula            -
    '- dblVariables - contains the variables being sent to the  -
    '- shapes formula                                           -
    '- intSelectedFormula - the selected formula                -
    '------------------------------------------------------------
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        Dim intSelectedFormula As Integer = lstFormulas.SelectedIndex

        Dim dblAnswer As Double
        If IsNumeric(txtFirstVariable.Text) Then
            If IsNumeric(txtSecndVariable.Text) Then
                If IsNumeric(txtThirdVariable.Text) Then
                    Dim dblVariables = New Double() {CDbl(txtFirstVariable.Text),
                                      CDbl(txtSecndVariable.Text),
                                      CDbl(txtThirdVariable.Text)}
                    Select Case intSelectedFormula
                        Case formula.PerimeterCicumferenceVolume
                            dblAnswer = lstShapes.SelectedItem.funPerimeterCircumferenceVolume(dblVariables)
                        Case formula.AreaSurfaceArea
                            dblAnswer = lstShapes.SelectedItem.funAreaSurfaceArea(dblVariables)
                        Case Else
                            Debug.WriteLine("Formula Selected was not one of the available options.")
                    End Select
                    txtAnswer.Text = dblAnswer.ToString()
                Else
                    'Third Variable is non numeric
                    MessageBox.Show("Variable must be numeric.")
                    txtThirdVariable.Text = "0"
                End If
            Else
                'Second variable is non numeric
                MessageBox.Show("Variable must be numeric.")
                txtSecndVariable.Text = "0"
            End If
        Else
            'First variable is non numeric
            MessageBox.Show("Variable must be numeric.")
            txtFirstVariable.Text = "0"
        End If

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnC_Click     -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when pressing the C button     -
    '- it will clear all the variable text boxes                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):              -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- txtBoxes - contains a list of all the text boxes         -
    '- txtFirstVariable, txt,SecondVaraible, txtThirdVariable,  -
    '- txtAnswer                                                -
    '------------------------------------------------------------
    Private Sub btnC_Click(sender As Object, e As EventArgs) Handles btnC.Click
        Dim txtBoxes = From txtBox In Me.Controls
                       Where TypeOf (txtBox) Is TextBox
                       Select txtBox
        For Each txtBox As TextBox In txtBoxes
            txtBox.Clear()
            txtBox.Text = "0"
        Next

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: btnCE_Click               -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when pressing the CE button    -
    '- it will clear the most recent text box                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- txtBoxes - contains a list of all the text boxes         -
    '- txtFirstVariable, txt,SecondVaraible, txtThirdVariable,  -
    '- txtAnswer                                                -
    '------------------------------------------------------------
    Private Sub btnCE_Click(sender As Object, e As EventArgs) Handles btnCE.Click
        If Not IsNothing(txtSelectedTextBox) Then
            txtSelectedTextBox.Clear()
            txtSelectedTextBox.Text = "0"
        Else
            MessageBox.Show("Please select a text box.")
        End If

    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: txtBox_GotFocus          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when a text box gains focus,   -
    '- it will set teh class variable txtSelectedTextBox        -
    '- to point to the selected text box, providing a way to    -
    '- remember the last text box                               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub txtBox_GotFocus(sender As TextBox, e As EventArgs) Handles txtFirstVariable.GotFocus,
                                                                           txtSecndVariable.GotFocus,
                                                                           txtThirdVariable.GotFocus
        'Sets the pointer for class variable to be the object(textbox) that called
        'the function
        txtSelectedTextBox = sender
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: frmChild_Load             -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when a frmChild is created     -
    '- it will create a list of shapes and then add the shapes  -
    '- to the txtShpae box                                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- lstUdtShapes - creates a list of clsShape objects        -
    '------------------------------------------------------------
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
    '------------------------------------------------------------
    '- Subprogram Name: lstFormulas_SelectedIndexChanged        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the index of lstFormula   -
    '- is changed. When the index changes the text boxes shown  -
    '- and the labels for the text boxes.                       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intNumVariables - the number of variables needed for the -
    '- shape to perform it's formulas                           -
    '------------------------------------------------------------
    Private Sub lstFormulas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles _
                                                                            lstFormulas.SelectedIndexChanged
        Dim intNumVariables As Integer = lstShapes.SelectedItem.GetstrMeasurementVariables().Count()

        'Prevents the user from selecting an index that is out of bounds
        'a -1 index is selected when a value is not chosen
        If lstFormulas.SelectedIndex >= 0 Then
            'Calculate only clickable after a formula is slected
            btnCalculate.Enabled = True
            With lstShapes.SelectedItem
                'Show only the number of text boxes needed for the shape
                'Label text will bet set to what teh  necessary varaible corresponds to
                Select Case intNumVariables
                    Case 1
                        lblFirstVariable.Visible = True
                        lblFirstVariable.Text = .GetstrMeasurementVariables()(0)
                        txtFirstVariable.Visible = True
                    Case 2
                        lblFirstVariable.Visible = True
                        lblFirstVariable.Text = .GetstrMeasurementVariables()(0)
                        txtFirstVariable.Visible = True
                        lblSecondVariable.Visible = True
                        lblSecondVariable.Text = .GetstrMeasurementVariables()(1)
                        txtSecndVariable.Visible = True
                    Case 3
                        lblFirstVariable.Visible = True
                        lblFirstVariable.Text = .GetstrMeasurementVariables()(0)
                        txtFirstVariable.Visible = True
                        lblSecondVariable.Visible = True
                        lblSecondVariable.Text = .GetstrMeasurementVariables()(1)
                        txtSecndVariable.Visible = True
                        lblThirdVariable.Visible = True
                        lblThirdVariable.Text = .GetstrMeasurementVariables()(2)
                        txtThirdVariable.Visible = True
                    Case Else
                        'For testing purposes, should not happen during normal operation
                        'of teh program.
                        Debug.WriteLine("Error occured when showing the txtVariable boxes")
                End Select
            End With
        End If


    End Sub
    '------------------------------------------------------------
    '- Subprogram Name: lstShapes_SelectedIndexChanged          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the index of lstShapes    -
    '- is changed. When the index changes it will hide the      -
    '- variable textboxes, labels, and disable the calculate    -
    '- button. It will also change the image shown to that of   -
    '- the selected image.                                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub lstShapes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstShapes.SelectedIndexChanged
        lstFormulas.Items.Clear()
        subHideVariables()
        With lstShapes.SelectedItem
            picShape.Image = Image.FromFile(.GetstrImagePath())
            For Each s In .GetstrFormulaTypes()
                lstFormulas.Items.Add(s)
            Next
        End With
    End Sub
    '------------------------------------------------------------
    '- Subprogram Name: subHideVariables                        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine hides the varaible text boxes and labels -
    '- as well as disabling teh btnCalculate                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub subHideVariables()
        btnCalculate.Enabled = False
        lblFirstVariable.Visible = False
        txtFirstVariable.Visible = False
        txtFirstVariable.Text = "0"
        lblSecondVariable.Visible = False
        txtSecndVariable.Visible = False
        txtSecndVariable.Text = "0"
        lblThirdVariable.Visible = False
        txtThirdVariable.Visible = False
        txtThirdVariable.Text = "0"
        txtAnswer.Text = ""
    End Sub
    '------------------------------------------------------------
    '- Subprogram Name: rdoMetric_CheckedChanged                -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will convert the answer to imperial or   -
    '- metric depending on the radial button selected           -
    '- Depending on the selected formula will determing the     -
    '- conversion factor used                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub rdoMetric_CheckedChanged(sender As Object, e As EventArgs) Handles rdoMetric.CheckedChanged
        If txtAnswer.Text <> "" Then
            If rdoMetric.Checked Then
                Select Case lstFormulas.SelectedItem
                    Case "Perimeter", "Circumference"
                        txtAnswer.Text = dblConvertCmToInch(txtAnswer.Text)
                    Case "Area", "Surface Area"
                        txtAnswer.Text = dblConvertCm2ToInch2(txtAnswer.Text)
                    Case "Volume"
                        txtAnswer.Text = dblConvertCm3ToInch3(txtAnswer.Text)
                    Case Else
                        Debug.WriteLine("Error changin metrics")
                End Select
            End If
        End If
    End Sub
    '------------------------------------------------------------
    '- Subprogram Name: rdoMetric_CheckedChanged                -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will convert the answer to imperial or   -
    '- metric depending on the radial button selected           -
    '- Depending on the selected formula will determing the     -
    '- conversion factor used                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub rdoUs_CheckedChanged(sender As Object, e As EventArgs) Handles rdoUs.CheckedChanged
        If txtAnswer.Text <> "" Then
            If rdoUs.Checked Then
                Select Case lstFormulas.SelectedItem
                    Case "Perimeter", "Circumference"
                        txtAnswer.Text = dblConvertInchToCm(txtAnswer.Text)
                    Case "Area", "Surface Area"
                        txtAnswer.Text = dblConvertInch2ToCm2(txtAnswer.Text)
                    Case "Volume"
                        txtAnswer.Text = dblConvertInch3ToCm3(txtAnswer.Text)
                    Case Else
                        Debug.WriteLine("Error changin metrics")
                End Select
            End If

        End If

    End Sub
    '------------------------------------------------------------
    '- Subprogram Name: frmChild_Closing                        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine will trigger on form closing if the      -
    '- answer field is "" then the form will close if the       -
    '- txtAnswer field is <> "" then prompt the user for exit   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender - the object calling the event handler            -
    '- e - event arguments                                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub frmChild_Closing(sender As frmChild, e As CancelEventArgs) Handles Me.Closing
        If txtAnswer.Text <> "" Then
            If MessageBox.Show("Are you sure you want to quit?", sender.Text, MessageBoxButtons.YesNo) _
                = DialogResult.No Then
                e.Cancel = True
            End If

        End If
    End Sub

#Enable Warning IDE1006 ' Naming Styles

End Class