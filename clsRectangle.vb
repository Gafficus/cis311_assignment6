'------------------------------------------------------------
'-                File Name : clsSquare                     - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 20 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the implementation of clsShape for    -
'- a rectangle                                              -
'------------------------------------------------------------	
Public Class clsRectangle : Inherits clsShape
    '------------------------------------------------------------
    '-                Subprogram Name: New                      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called on the creation of a new       -
    '- rectangle it will set the name, variable names,          -
    '- and formula names for the rectangle                      -    
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Public Sub New()
        MyBase.SetstrName("2D - Rectangle")
        MyBase.SetstrImagePath("rectangle.jpg")
        MyBase.GetstrMeasurementVariables.Add("Length")
        MyBase.GetstrMeasurementVariables.Add("Width")
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
    '- for finding the Perimeter of a rectangle                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 -length                                                -
    '- 1 - width                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dbLength - contains the length of the rectangle          -
    '- dblWidth - contains the width of the rectangle           -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter of the rectangle                  -
    '- 2(lenght + width)                                        -
    '------------------------------------------------------------
    Public Overrides Function funPerimeterCircumferenceVolume(ByVal ParamArray variables() As Double) As Double
        Dim dblLength As Double = variables(0)
        Dim dblWidth As Double = variables(1)
        Return 2 * (dblLength + dblWidth)
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
    '- for finding the area of a rectangle                      -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function. In this case it contains                       -
    '- 0 - length                                               -
    '- 1 - width                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- dbLength - contains the length of the rectangle          -
    '- dblWidth - contains the width of the rectangle           -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter of the rectangle                  -
    '- length * width                                           -
    '------------------------------------------------------------
    Public Overrides Function funAreaSurfaceArea(ByVal ParamArray variables() As Double) As Double
        Dim dblLength As Double = variables(0)
        Dim dblWidth As Double = variables(1)
        Return dblLength * dblWidth
    End Function
End Class
