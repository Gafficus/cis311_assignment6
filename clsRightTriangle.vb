'------------------------------------------------------------
'-                File Name : clsRightTriangle              - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 20 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a right angle triangle.                                  -
'------------------------------------------------------------	
Public Class clsRightTriangle : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New            -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is on creation of a clsRightTriangle     -
    '- object instanteation. It will set the name of the obj    -
    '- to the proper string and set the string of variables     -
    '- to contain a Base and a Height, it will also set the     -
    '- name of the formula to Perimeter and Area                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New()
        MyBase.SetstrName("2D - Right Triangle")
        MyBase.GetstrMeasurementVariables.Add("Base")
        MyBase.GetstrMeasurementVariables.Add("Height")
        MyBase.GetstrFormulaTypes.Add("Perimeter")
        MyBase.GetstrFormulaTypes.Add("Area")
    End Sub
    '------------------------------------------------------------
    '-    Subprogram Name: funPerimeterCircumferenceVolume      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine provides the implementation from clsShape–
    '- for finding the Perimeter of a right angle triangle      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Base                                                 -
    '- 1 - Height
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblBase - contains the Base value of the triangle        -
    '- dblheight - contains the height of the trianlgle         -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter of the right triangle             -
    '- height + base + (height^2 + base^2)                     -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblBase As Double = variables(0)
        Dim dblHeight As Double = variables(1)
        Return dblHeight + dblBase + Math.Sqrt(Math.Pow(dblHeight, 2) + Math.Pow(dblBase, 2))
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: funAreaSurfaceArea       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function will calculate the area of a right angle   –
    '- triangle.                                                -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- variables() - contains the varaibles being sent.         -
    '- 0 - Base
    '- 1 - Height
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblBase - Holds the base value of the triangle           -
    '- dblHeight - Hold the height of the triangle              -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the area of the triangle                        -
    '- (1/2) * base * height                                    -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblBase As Double = variables(0)
        Dim dblHeight As Double = variables(1)
        Return 0.5 * dblBase * dblHeight
    End Function
End Class