/*namespace ColumnDesign.Modules
{
    public class CalcWeight_Function
    {
        'Calculates weight of 1 given column from ply, studs, clamps, and braces (only 7-11 and 11-19 braces).
'Then 50 lbs is added and that total is rounded up to the nearest 100 lbs.

Option Explicit
    'Define weights, public constants may be used anywhere
    Public Const wt_ply As Double = 2.2         'psf
    Public Const wt_2x4 As Double = 1.31        'lb/ft
    Public Const wt_LVL As Double = 1.8        'lb/ft ????????????????????????????????????? what is LVL linear weight
    Public Const wt_08_24_clamp As Double = 76  'lbs
    Public Const wt_12_36_clamp As Double = 100 'lbs
    Public Const wt_24_48_clamp As Double = 123 'lbs
    Public Const wt_36_scissor_clamp As Double = 40 'lbs
    Public Const wt_48_scissor_clamp As Double = 56 'lbs
    Public Const wt_60_scissor_clamp As Double = 85 'lbs
    Public Const wt_07_11_brace As Double = 52  'lbs
    Public Const wt_11_19_brace As Double = 76  'lbs
    Public Const wt_extra As Double = 50        'lbs
    
    
Function wt_total(x, y, z, n_studs_x, n_studs_y, stud_type, n_clamps, clamp_size, brace_size) As Double

    Dim wt_stud As Double
    Dim wt_clamp As Double
    Dim wt_brace As Double
    
'Pick weights based on stud type, clamp size, and brace size
    Select Case stud_type
        Case Is = 1
            wt_stud = wt_2x4
        Case Is = 2
            wt_stud = wt_LVL
        Case Else
            MsgBox ("Error: Invalid stud type detected in weight calculation")
    End Select
    
    Select Case clamp_size
        Case Is = 1
            wt_clamp = wt_08_24_clamp
        Case Is = 2
            wt_clamp = wt_12_36_clamp
        Case Is = 3
            wt_clamp = wt_24_48_clamp
        Case Is = 4
            wt_clamp = wt_36_scissor_clamp
        Case Is = 5
            wt_clamp = wt_48_scissor_clamp
        Case Is = 6
            wt_clamp = wt_60_scissor_clamp
        Case Else
            MsgBox ("Error: Invalid clamp size detected in weight calculation")
    End Select
    
    Select Case brace_size
        Case Is = 0
            wt_brace = 0
        Case Is = 1
            wt_brace = wt_07_11_brace
        Case Is = 2
            wt_brace = wt_11_19_brace
        Case Else
            MsgBox ("Error: Invalid brace size detected in weight calculation")
    End Select
    
    'Sum weights
    wt_total = _
    wt_ply * (((x + 2.25) * z * 2 / 144) + ((y + 2.25) * z * 2 / 144)) + _
    wt_stud * (z / 12) * 2 * (n_studs_x + n_studs_y) + _
    wt_clamp * n_clamps + _
    wt_brace * 3
    
    'Add 50 lbs and round up to the nearest 100
    wt_total = ColumnCreator.RoundUp((wt_total + wt_extra) / 100)
    wt_total = wt_total * 100
    
End Function


Function wt_panel(x, y, z, n_studs_x, n_studs_y, stud_type) As Double
'This function is for calculating the weight of one side of the column form.
    Dim wt_stud As Double
    Dim larger_side As Double
    Dim most_studs As Double
    
'Pick weights based on stud type, clamp size, and brace size
    Select Case stud_type
        Case Is = 1
            wt_stud = wt_2x4
        Case Is = 2
            wt_stud = wt_LVL
        Case Else
            MsgBox ("Error: Invalid stud type detected in weight calculation")
    End Select
    
    'Find if x or y side is larger
    If x > y Then
        larger_side = x
    Else
        larger_side = y
    End If
    
    'Find out if the x or y side has more studs
    If n_studs_x > n_studs_y Then
        most_studs = n_studs_x
    Else
        most_studs = n_studs_y
    End If
    
    'Sum weights
    wt_panel = _
    wt_ply * ((larger_side + 4.5) * z / 144) + _
    wt_stud * (z / 12) * most_studs
    
    'Increase weight by 10% and round it up to the nearest 10 pounds
    wt_panel = ColumnCreator.RoundUp((wt_panel * 1.1) / 10)
    wt_panel = wt_panel * 10

    
    
End Function


    }
}*/