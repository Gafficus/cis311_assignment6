'------------------------------------------------------------
'-                File Name : clsCircle              - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 23 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a circle                                                 -
'------------------------------------------------------------	
Public Class clsCircle : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New                      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is on creation of a cls Circle           -
    '- object instanteation. It will set the name of the obj    -
    '- to the proper string and set the string of variables     -
    '- to contain a radius and set the name of the formula to   -
    '-circumference and Area                                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New()
        MyBase.SetstrName("2D - Circle")
        MyBase.SetstrImagePath("circle.jpg")
        MyBase.GetstrMeasurementVariables.Add("Radius")
        MyBase.GetstrFormulaTypes.Add("Circumference")
        MyBase.GetstrFormulaTypes.Add("Area")
    End Sub
    '------------------------------------------------------------
    '-    Subprogram Name: funPerimeterCircumferenceVolume      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine provides the implementation from clsShape–
    '- for finding the Perimeter of a circle                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Radius                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - contains the radius value of the circle      -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the circumference of the circle                 -
    '- 2 * pi * radius                                          -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Return 2 * Math.PI * dblRadius
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: funAreaSurfaceArea       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function will calculate the area of a circle        -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- variables() - contains the varaibles being sent.         -
    '- 0 - Radius                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - Holds the radius of the circle               -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the area of the circle                          -
    '- pi * r^2                                                 -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Return Math.PI * Math.Pow(dblRadius, 2)
    End Function
End Class