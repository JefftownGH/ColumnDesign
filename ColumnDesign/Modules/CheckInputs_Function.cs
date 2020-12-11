/*namespace ColumnDesign.Modules
{
    public class CheckInputs_Function
    {
        Public Function CheckInputs(ByVal x, ByVal y, ByVal z, ByVal qty As Double, qtyCheck As Integer) As Integer
    'Check that all inputs are within defined bounds NOTE: Change this to 2read limits from the data files
    'qtyCheck = 0 means the quantity box is not considered
    
    Dim min_width As Integer
    Dim min_length As Integer
    Dim min_height As Integer
    Dim max_width As Integer
    Dim max_length As Integer
    Dim max_height As Integer
        
    If ColumnCreator.MultiPage1.Value = 0 Then
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'                    GATES COLUMNS
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #

        'Define maximum column dimensions
        min_width = 12
        min_length = 12
        min_height = 96
        max_width = 48
        max_length = 48
        max_height = 240
        
        'Assume checks will pass
        CheckInputs = 1
        
        'If any check fails it will not pass
        If Round(x, 3) < min_width Then
            CheckInputs = 0
        ElseIf Round(x, 3) > max_width Then
            CheckInputs = 0
        End If
    
        If Round(y, 3) < min_length Then
            CheckInputs = 0
        ElseIf Round(y, 3) > max_length Then
            CheckInputs = 0
        End If
        
        If Round(z, 3) < min_height Then
            CheckInputs = 0
        ElseIf Round(z, 3) > max_height Then
            CheckInputs = 0
        End If
        
        If qtyCheck = 1 And qty = 0 Then
            CheckInputs = 0
        End If
        
        'Control the background colors to indicate when an entered value is out of bounds (without passing or failing the check)

        If Round(x, 3) < min_width And x <> 0 Then
            ColumnCreator.WidthBox.BackColor = &H8080FF
        ElseIf Round(x, 3) > max_width And x <> 0 Then
            ColumnCreator.WidthBox.BackColor = &H8080FF
        Else
            ColumnCreator.WidthBox.BackColor = &H80000005
        End If
        
        If Round(y, 3) < min_length And y <> 0 Then
            ColumnCreator.LengthBox.BackColor = &H8080FF
        ElseIf Round(y, 3) > max_length And y <> 0 Then
            ColumnCreator.LengthBox.BackColor = &H8080FF
        Else
            ColumnCreator.LengthBox.BackColor = &H80000005
        End If
        
        If Round(z, 3) < min_height And z <> 0 Then
            ColumnCreator.HeightBox.BackColor = &H8080FF
        ElseIf Round(z, 3) > max_height And z <> 0 Then
            ColumnCreator.HeightBox.BackColor = &H8080FF
        Else
            ColumnCreator.HeightBox.BackColor = &H80000005
        End If
        
        
        
    
    ElseIf ColumnCreator.MultiPage1.Value = 1 Then
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'                 SCISSOR CLAMP COLUMNS
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #

        'Define maximum column dimensions
        min_width = 8
        min_length = 8
        min_height = 12
        max_width = 46
        max_length = 46
        max_height = 480 ' 40'-0" (10 4' tall plywood sheets is 1 under the plywood drawing limit)
        
        'Assume checks will pass
        CheckInputs = 1
        
        'If any check fails it will not pass
        If Round(x, 3) < min_width Then
            CheckInputs = 0
            ColumnCreator.slblMaxHt.Caption = "N/A"
        ElseIf Round(x, 3) > max_width Then
            CheckInputs = 0
        End If
    
        If Round(y, 3) < min_length Then
            CheckInputs = 0
            ColumnCreator.slblMaxHt.Caption = "N/A"
        ElseIf Round(y, 3) > max_length Then
            CheckInputs = 0
        End If
        
        If Round(z, 3) < min_height Then
            CheckInputs = 0
        ElseIf Round(z, 3) > max_height Then
            CheckInputs = 0
        End If
        
        If qtyCheck = 1 And qty = 0 Then
            CheckInputs = 0
        End If
        
        'Control the background colors to indicate when an entered value is out of bounds (without passing or failing the check)

        If Round(x, 3) < min_width And x <> 0 Then
            ColumnCreator.sWidthBox.BackColor = &H8080FF
        ElseIf Round(x, 3) > max_width And x <> 0 Then
            ColumnCreator.sWidthBox.BackColor = &H8080FF
        Else
            ColumnCreator.sWidthBox.BackColor = &H80000005
        End If
        
        If Round(y, 3) < min_length And y <> 0 Then
            ColumnCreator.sLengthBox.BackColor = &H8080FF
        ElseIf Round(y, 3) > max_length And y <> 0 Then
            ColumnCreator.sLengthBox.BackColor = &H8080FF
        Else
            ColumnCreator.sLengthBox.BackColor = &H80000005
        End If
        
        If Round(z, 3) < min_height And z <> 0 Then
            ColumnCreator.sHeightBox.BackColor = &H8080FF
        ElseIf Round(z, 3) > max_height And z <> 0 Then
            ColumnCreator.sHeightBox.BackColor = &H8080FF
        Else
            ColumnCreator.sHeightBox.BackColor = &H80000005
        End If
        
    End If
End Function

    }
}*/