Public MustInherit Class clsShape
    Private _strName As String

    Private Property strName() As String
        Get
            Return _strName
        End Get
        Set
            _strName = Value
        End Set
    End Property

    Public MustOverride Function funPerimeterCircumferenceVolume(dblOne As Double, dblTwo As Double) As Double
    Public MustOverride Function funAreaSurfaceArea(dblOne As Double, dblTwo As Double) As Double
End Class
