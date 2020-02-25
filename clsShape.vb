'------------------------------------------------------------
'-                File Name : clsShape                     - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 20 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the abstract definition of a shpae    -
'- class.
'------------------------------------------------------------	
Public MustInherit Class clsShape
    Private _strMeasurementVariables As New List(Of String)
    Private _strFormulaTypes As New List(Of String)
    Private _strName As String
    '------------------------------------------------------
    '- BEGIN GET AND SET FUNCTIONS
    Public Function GetstrFormulaTypes() As List(Of String)
        Return _strFormulaTypes
    End Function

    Public Sub SetstrFormulaTypes(AutoPropertyValue As List(Of String))
        _strFormulaTypes = AutoPropertyValue
    End Sub

    Public Function GetstrMeasurementVariables() As List(Of String)
        Return _strMeasurementVariables
    End Function

    Public Sub SetstrMeasurementVariables(AutoPropertyValue As List(Of String))
        _strMeasurementVariables = AutoPropertyValue
    End Sub


    Public Function GetstrName() As String
        Return _strName
    End Function

    Public Sub SetstrName(AutoPropertyValue As String)
        _strName = AutoPropertyValue
    End Sub
    '- END GET AND SET FUNCTIONS
    '------------------------------------------------------------

    '------------------------------------------------------------
    '-    Subprogram Name: funPerimeterCircumferenceVolume      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine provides the abstract method for clsShape-
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function.                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the perimeter, circumference, or volume of the  -
    '- shape                                                    -
    '------------------------------------------------------------
    Public MustOverride Function funPerimeterCircumferenceVolume(ByVal ParamArray variables() As Double) As Double
    '------------------------------------------------------------
    '-    Subprogram Name: funAreaSurfaceArea      -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 20 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '- This subroutine provides abstract method for clsShape    -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- vairables()- contains the variables being sent to the    -
    '- function.                                                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (none)                                                   -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the area or surface area of the shape           -
    '------------------------------------------------------------
    Public MustOverride Function funAreaSurfaceArea(ByVal ParamArray variables() As Double) As Double

    Public Overrides Function ToString() As String
        Return _strName
    End Function
End Class
