/*namespace ColumnDesign.Modules
{
    public class ConvertNumberToFeetInches
    {
        
Function ConvertFtIn(ByVal num) As String

    '================================================
    'This version properly converts negative numbers
    '================================================
    
    Dim xFt As String
    Dim xIn As String
    Dim xFra As String
    Dim strNeg As String
    
    If num = "" Or num = " " Then
        ConvertFtIn = 0
        Exit Function
    End If
    
    'If the number is negative, make it positive, then multiply add a "-" character at the very end
    If num < 0 Then
        strNeg = "-"
        num = num * -1
    Else
        strNeg = ""
    End If
    
    'Add a hair then round DOWN to nearest 1/8th inch
    'num = (Round((num + 0.01) * 8 + 0.5, 0) - 1) / 8
    
    'Round to the nearest 1/8th inch
    num = (Round((num * 8) - 0.000000001, 0)) / 8
    
    xFt = CStr(Int(Round(num / 12, 3)))
    xIn = CStr(num - 12 * Int(Round(num / 12, 3)))
    
    'If inches are not an integer convert decimal to fraction
    If Round(num - Int(Round(num, 3)), 3) <> 0 Then
        xFra = Dec2Frac(12 * ((num / 12) - Int(num / 12)) - Int(num - 12 * Int(num / 12))) 'Inches minus integer inches
        xIn = Int(num - 12 * Int(num / 12)) 'Redefine inches without the fraction
    End If
    
    'If ft = 0 as a string don't print ft
    If xFt = "0" Then
        If Round(num - Int(Round(num, 3)), 3) <> 0 Then
            ConvertFtIn = strNeg & xIn & Space(1) & xFra & """"  'Ft = 0 and inches are fractional
        Else
            ConvertFtIn = strNeg & xIn & """"    'Ft = 0 and inches are whole integers
        End If
    Else
        If Round(num - Int(Round(num, 3)), 3) <> 0 Then
            ConvertFtIn = strNeg & xFt & "'-" & xIn & Space(1) & xFra & """" 'Ft <> 0 and inches are fractional
        Else
            ConvertFtIn = strNeg & xFt & "'-" & xIn & """"   'Ft <> 0 and inches are whole integers
        End If
    End If
    
End Function

Public Function Dec2Frac(ByVal dFraction As Double) As String
    
    'Type: Public Function
    'Name: Decimal to Fraction Converter
    'Purpose: Converts a passed decimal to a fraction
    'Limitations: Will not convert to mixed fractions
    'Author: Unknown
    'Arguments: dFraction is the decimal to convert
    'Return Value: String representingthe fraction
    'Useage: Dec2Frac(.125)
    'Notes:
    
    Dim df As Double
    Dim lUpperPart As Long
    Dim lLowerPart As Long
    
    lUpperPart = 1
    lLowerPart = 1
    
    dFraction = Round(dFraction, 3) 'Cleans up messy double values
    df = lUpperPart / lLowerPart
    
    While (df <> dFraction)
        If (df < dFraction) Then
            lUpperPart = lUpperPart + 1
        Else
            lLowerPart = lLowerPart + 1
            lUpperPart = dFraction * lLowerPart
        End If
        df = lUpperPart / lLowerPart
    Wend
    
    Dec2Frac = CStr(lUpperPart) & "/" & CStr(lLowerPart)

End Function

    }
}*/