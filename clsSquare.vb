'------------------------------------------------------------
'-                File Name : clsSquare                     - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 20 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a square                                                 -
'------------------------------------------------------------	
Public Class clsSquare : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New                      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called on the creation of a new square-
    '- it will set the name, variable names, and formula names  -
    '- for the square.
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New()
        MyBase.SetstrName("2D - Square")
        MyBase.SetstrImagePath("square.jpg")
        MyBase.GetstrMeasurementVariables.Add("Length")
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
    '- for finding the Perimeter of a square                    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 -length                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dbLength - contains the length of the square             -   
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter of the square                     -
    '- length+length+length+length                              -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ParamArray variables() As Double) As Double
        Dim dblLength As Double = variables(0)
        Return 4 * dblLength
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: funAreaSurfaceArea                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine provides the implementation from clsShape–
    '- for finding the area of a square                         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 -length                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dbLength - contains the length of the square             -   
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter of the square                     -
    '- length^2                                                 -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ParamArray variables() As Double) As Double
        Dim dblLength As Double = variables(0)
        Return Math.Pow(dblLength, 2)
    End Function
End Class