﻿'------------------------------------------------------------
'-                File Name : clsCone                       - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 25 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a cone                                                   -
'------------------------------------------------------------	
Public Class clsCone : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New                      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is on creation of a cls cone             -
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
        MyBase.SetstrName("3D - Cone")
        MyBase.SetstrImagePath("cone.jpg")
        MyBase.GetstrMeasurementVariables.Add("Radius")
        MyBase.GetstrMeasurementVariables.Add("Height")
        MyBase.GetstrFormulaTypes.Add("Volume")
        MyBase.GetstrFormulaTypes.Add("Surface Area")
    End Sub
    '------------------------------------------------------------
    '-    Subprogram Name: funPerimeterCircumferenceVolume      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine provides the implementation from clsShape–
    '- for finding the Perimeter of a cone                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Radius                                               -
    '- 0 - Height                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - hold the value of the radius                 -
    '- dblheight - hold the value of the height                 -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the Volume of the cone                          -
    '- 1/3 * pi *  r^2 * h                                      -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Dim dblheight As Double = variables(1)
        Return (1 * Math.PI * Math.Pow(dblRadius, 2) * dblheight) / (3)
    End Function
    '------------------------------------------------------------
    '-                Subprogram Name: funAreaSurfaceArea       -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This function will calculate the surface area of a       -
    '- sphere                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - Radius                                               -
    '- 1 - Height                                               -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dblRadius - holds the radius value of the sphere         -
    '- dblheight - holds the height value of the sphere         -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the surface area of the sphere                  -
    '- (PI * R) * (R +SQRT(R^2 + H^2))                          -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblRadius As Double = variables(0)
        Dim dblHeight As Double = variables(1)
        Return (Math.PI * dblRadius) * (dblRadius + Math.Sqrt(Math.Pow(dblRadius, 2) + Math.Pow(dblHeight, 2)))
    End Function
End Class