'------------------------------------------------------------
'-                File Name : modGlobalConversion           - 
'-                Part of Project: Assign6                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 25 Feb 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains a psuedo static class to convert      -
'- numbers from imperial to metric and vice versa           -
'------------------------------------------------------------	
Module modGlobalConversion
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertInchToCm                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts an inch measurement to cm       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertInchToCm(ByVal number As Double)
        Return number * 2.54
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertCmToInch                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts an cm measurement to inch       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertCmToInch(ByVal number As Double)
        Return number * 0.393701
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertInch2ToCm2                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts an square inch measurement to   -
    ' square cm                                                 -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertInch2ToCm2(ByVal number As Double)
        Return number * 6.4516
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertCm2ToInch2                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts an square cm to square inch     -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertCm2ToInch2(ByVal number As Double)
        Return number * 0.155
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertInch3ToCm3                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts cubic inch to cubic cm          -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertInch3ToCm3(ByVal number As Double)
        Return number * 16.3871
    End Function
    '------------------------------------------------------------
    '-    Subprogram Name: dblConvertCm3ToInch3                   -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 25 Feb 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine converts an cubic cm to cubic inch       -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- number - the number to be converted                      -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '-(None)                                                    -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- Double - the converted number                            -
    '------------------------------------------------------------
    Public Function dblConvertCm3ToInch3(ByVal number As Double)
        Return number * 0.0610237
    End Function
End Module
