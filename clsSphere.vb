'------------------------------------------------------------
'-                File Name : clsSphere                     - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 24 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a sphere                                                 -
'------------------------------------------------------------	
Public Class clsSphere : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New                      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 24 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is on creation of a cls Sphere           -
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
        MyBase.SetstrName("3D - Sphere")
        MyBase.GetstrMeasurementVariables.Add("Radius")
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
    '- 0 - Radius                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - hold the value of the radius                 -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the Volume of the sphere                        -
    '- 4/3 * pi * r^3                                           -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Return (4 * Math.PI * Math.Pow(dblRadius, 3)) / (3)
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: funAreaSurfaceArea       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function will calculate the surface area of a sphere-
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Radius                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - holds the radius value of the sphere         -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the surface area of the sphere                  -
    '- 4 * pi * r^2                                             -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Return (4 * Math.PI * Math.Pow(dblRadius, 2))
    End Function
End Class