'------------------------------------------------------------
'-                File Name : clsCube                       - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 23 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a cube                                                   -
'------------------------------------------------------------	
Public Class clsCube : Inherits clsShape
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
    '-volume and Surface Area                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New()
        MyBase.SetstrName("3D - Cube")
        MyBase.SetstrImagePath("cube.jpg")
        MyBase.GetstrMeasurementVariables.Add("Height")
        MyBase.GetstrMeasurementVariables.Add("Width")
        MyBase.GetstrMeasurementVariables.Add("Length")
        MyBase.GetstrFormulaTypes.Add("Volume")
        MyBase.GetstrFormulaTypes.Add("Surface Area")
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
    '- for finding the Perimeter of a cube                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Height                                               -
    '- 1 - Width                                                -
    '- 2 - Length                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblHeight - holds the height of the cube                 -
    '- dblLength - holds the Length of the cube                 -
    '- dblWidth -  holds the width of the cube                  -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the Volume of the cube                          -
    '- height * width * length                                  -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblHeight As Double = variables(0)
        Dim dblWidth As Double = variables(1)
        Dim dblLength As Double = variables(2)
        Return dblHeight * dblLength * dblWidth
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: funAreaSurfaceArea       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function will calculate the surface area of a cube  -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Height                                               -
    '- 1 - Width                                                -
    '- 2 - Length                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblHeight - holds the height of the cube                 -
    '- dblLength - holds the Length of the cube                 -
    '- dblWidth -  holds the width of the cube                  -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the area of the circle                          -
    '- 2LW + 2LH + 2WH                                          -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblHeight As Double = variables(0)
        Dim dblWidth As Double = variables(1)
        Dim dblLength As Double = variables(2)
        Return (2 * dblLength * dblWidth) + (2 * dblLength * dblHeight) + (2 * dblWidth * dblHeight)
    End Function
End Class