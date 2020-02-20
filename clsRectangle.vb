Public Class clsRectangle : Inherits clsShape

    Public Sub New()
    End Sub

    Public Overrides Function funPerimeterCircumferenceVolume(dblOne As Double, dblTwo As Double) As Double
        Return 2 * (dblOne * dblTwo)
    End Function
    Public Overrides Function funAreaSurfaceArea(dblOne As Double, dblTwo As Double) As Double
        Throw New NotImplementedException()
    End Function
End Class
