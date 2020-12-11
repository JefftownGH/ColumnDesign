/*namespace ColumnDesign.Modules
{
    public class GetPlySeams_Function
    {
        Public Function GetPlySeams(ByVal x, ByVal y, ByVal z) As Double()
    Dim ply_width_x As Double
    Dim ply_width_y As Double
    Dim ply_top_ht As Double
    Dim ply_bot_n As Integer 'Number of 8'-0" by x'-x" sheets on the bottom of the column
    Dim ply_top_ht_min As Double: ply_top_ht_min = 6 'Smallest allowable strip of plywood at top of column
    Dim ply_seams_fun() As Double '1 dimensional matrix that stores vertical length of each ply piece
    Dim max_ply_ht As Double
    
    If ColumnCreator.MultiPage1.Value = 0 Then 'GATES COLUMNS

        'Add 1.5 inches to nominal width for ply sheet width
        ply_width_x = x + 1.5
        ply_width_y = y + 1.5
        
    ElseIf ColumnCreator.MultiPage1.Value = 1 Then 'SCISSOR CLAMP COLUMNS
        'x: Add 0 inches to nominal width for ply sheet width on the short side
        'y: Add 4.5 inches to nominal width for ply sheet width on the long side
        ply_width_x = x
        ply_width_y = y + 4.5
    End If
    
    'If column is very wide, use rotated sheets and thus 48" max height
    If Round(ply_width_x, 3) > 48 Or Round(ply_width_y, 3) > 48 Then
        max_ply_ht = 48
    Else
        max_ply_ht = 96
    End If
    
    'Assume sheets aligned vertically. How many of those fit on the bottom and what is the remainder in inches on top?
    ply_top_ht = z - max_ply_ht * Int(Round(z / max_ply_ht, 3)) 'Int will floor to the nearest integer
    ply_bot_n = Int(Round(z / max_ply_ht, 3))

    'Add the typical sheets plus 1 if a smaller top ply is used
    ReDim ply_seams_fun(1 To ply_bot_n + Abs(Round(ply_top_ht, 3) <> 0))

    'If only a single sheet is required, use that and jump to the end
    If Round(z, 3) <= max_ply_ht Then
        ply_seams_fun(1) = z
        GoTo PlySeamsObtained
    End If
    
    'Start filling ply_seams_fun array with typical 8'-0" or 4'-0" sheets and fill the last value with the height of the top sheet, if any
    For i = 1 To ply_bot_n
        ply_seams_fun(i) = max_ply_ht
    Next i
    If Round(ply_top_ht, 3) <> 0 Then
        ply_seams_fun(UBound(ply_seams_fun)) = ply_top_ht
    End If
    
    'If top piece is too small and nonzero add 48" to it and use a 48" piece below it.
    If Round(ply_top_ht, 3) < Round(ply_top_ht_min, 3) And Round(ply_top_ht, 3) > 0 Then
        ply_top_ht = ply_top_ht + 48 * max_ply_ht / 96 '1/2 of max ply ht (24" or 48")
        ply_bot_n = Int(Round((z - 48 * max_ply_ht / 96 - ply_top_ht) / max_ply_ht, 3)) 'Recompute number of 8' or 4' ply pieces on the bottom, less the new top pieces
        ReDim ply_seams_fun(1 To ply_bot_n + 2) 'Delete old data in this array, resize it for the new ply arrangement
        
        For i = 1 To ply_bot_n
            ply_seams_fun(i) = max_ply_ht
        Next i
        ply_seams_fun(UBound(ply_seams_fun) - 1) = 48 * max_ply_ht / 96
        ply_seams_fun(UBound(ply_seams_fun)) = ply_top_ht
    End If
PlySeamsObtained:

    GetPlySeams = ply_seams_fun

End Function

    }
}*/