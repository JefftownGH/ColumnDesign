/*namespace ColumnDesign.Modules
{
    public class UpdatePly
    {
        Public Function UpdatePly(InputSource As Integer)

' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'                  FOR GATES COLUMNS
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #

    'Input source of 0 is for changes from the x, y, and z boxes or from changing the axis
    'Input source of 1 is for changes from boxPlySeams
    'Input source of 2 is for changes from the window checkboxes or the window position dimensions
    'Input source of 3 is for changes from the z box only

    Dim x As Double
    Dim y As Double
    Dim z As Double

    x = ConvertToNum(ColumnCreator.WidthBox)
    y = ConvertToNum(ColumnCreator.LengthBox)
    z = ConvertToNum(ColumnCreator.HeightBox)

    'Detect very large values. If found, just abort the function
    If x > 99 Or y > 99 Or z > 999 Then
        Exit Function
    End If
    
    'Build collection so line images can be referenced by their index
    Set Coll = Collection_Function.BuildCollection()

    'If just changing the height (z), recompute pour window location
    If InputSource = 3 Then
        ColumnCreator.WinDim1.Value = ConvertFtIn(Round(z / 5, 0))
        ColumnCreator.WinDim2.Value = ConvertFtIn(z - ConvertToNum(ColumnCreator.WinDim1.Value))
        'GoTo UpdatePlyColorCheck
    End If
            
            
    'Clear visuals
    'Window dimensions are handled separately because if they are hidden at this point they will get hidden and unhidden under normal usage, which removes focus from the dimension, this makes it difficult to type in dimensions.
    ColumnCreator.img_line_1.Visible = False
    ColumnCreator.img_line_2.Visible = False
    ColumnCreator.img_line_3.Visible = False
    ColumnCreator.img_line_win.Visible = False
    For i = 1 To Coll.Count
        Coll.Item(i).Visible = False
    Next i
    
    'Triggers only when x, y, and z are all provided
    If Round(x * y * z, 3) <> 0 Then
    
        'Check that x, y, and z inputs are within a valid range. If not, hide any window dimensions and skip the rest of this function. (qtyCheck = 0 means the quantity box is not checked)
        If CheckInputs(ConvertToNum(ColumnCreator.WidthBox.Value), ConvertToNum(ColumnCreator.LengthBox.Value), ConvertToNum(ColumnCreator.HeightBox.Value), ConvertToNum(ColumnCreator.QuantityBox.Value), 0) = 0 Then
            ColumnCreator.WinDim1.Visible = False
            ColumnCreator.WinDim2.Visible = False
            GoTo UpdatePlyLastLine
        End If
    
        'Enable drawing buttons (may be disabled later if ply seams are invalid)
        Call EnableDrawingButtons
        
'===============================================================================================================================
'   Manage userform objects and ply_seams textbox
'===============================================================================================================================

        'Get plywood seams layout
        Dim ply_seams() As Double
        Dim temp_ht As Double: temp_ht = 0
        
        If InputSource = 3 Or ColumnCreator.boxPlySeams.Value = "" Then
            'Generate the ply seams
            ply_seams = GetPlySeams(x, y, z)
            
            'Hide any previous error messages because these fresh ply seams should be flawless
            ColumnCreator.txtPlyError.Visible = False
        ElseIf InputSource = 0 Or InputSource = 1 Or InputSource = 2 Then
            'Read the ply seams
            ply_seams = ReadSizes(ColumnCreator.boxPlySeams.Value)
            
            'If ply seams are invalid: disable the drawing buttons, hide window dimensions, and skip the rest of this function
            If ValidatePlySeams(ply_seams, x, y, z) <> 1 Then
                Call DisableDrawingButtons
                ColumnCreator.WinDim1.Visible = False
                ColumnCreator.WinDim2.Visible = False
                GoTo UpdatePlyLastLine
            End If
            

        End If
        
        'If quantity is missing, disable drawing buttons and continue
        If ConvertToNum(ColumnCreator.QuantityBox.Value) = 0 Then
            Call DisableDrawingButtons
        End If
            
        'Populate the textbox with a new plywood layout
        If InputSource = 0 Or InputSource = 3 Then  '  Or InputSource = 2
            'If input is from the x or y box and there's already a ply layout, don't replace it
            If InputSource = 0 And ColumnCreator.boxPlySeams.Value <> "" Then
                GoTo SkipPlyUpdate
            End If
            
            Dim strPlySeams As String
            For i = 1 To UBound(ply_seams)
                If i <> UBound(ply_seams) Then
                    strPlySeams = strPlySeams & ConvertFtIn(ply_seams(i)) & ", "
                ElseIf i = UBound(ply_seams) Then
                    strPlySeams = strPlySeams & ConvertFtIn(ply_seams(i))
                End If
            Next i
            ColumnCreator.boxPlySeams.Value = strPlySeams
        End If
SkipPlyUpdate:

'===============================================================================================================================
'   Draw the column on the userform
'===============================================================================================================================
        'Define origin point for column sketch
        Dim pt_draw(0 To 2) As Double
        pt_draw(0) = 200
        pt_draw(1) = 216
        pt_draw(2) = 0
        
        'Define max pixel height and actual pixel width for x and y sides
        Dim px_z_max As Double
        Dim px_x As Double
        Dim px_y As Double
        px_z_max = 200
        px_x = px_z_max * (x / z)
        px_y = px_z_max * (y / z)
        
        'Change how this is defined for view types later
        If ColumnCreator.lblAxis.Caption = "X" Then
            px_width = px_x
        ElseIf ColumnCreator.lblAxis.Caption = "Y" Then
            px_width = px_y
        Else
            MsgBox ("Error: Horizontal dimension not x or y")
        End If
        
        'Draw "frame" of column: left, right, and top
        'Left
        ColumnCreator.img_line_1.Left = pt_draw(0) - px_width / 2
        ColumnCreator.img_line_1.top = pt_draw(1) - px_z_max
        ColumnCreator.img_line_1.Width = 2
        ColumnCreator.img_line_1.Height = px_z_max
        ColumnCreator.img_line_1.Visible = True
    
        'Right
        ColumnCreator.img_line_2.Left = pt_draw(0) + px_width / 2
        ColumnCreator.img_line_2.top = pt_draw(1) - px_z_max
        ColumnCreator.img_line_2.Width = 2
        ColumnCreator.img_line_2.Height = px_z_max
        ColumnCreator.img_line_2.Visible = True
    
        'Top
        ColumnCreator.img_line_3.Left = pt_draw(0) - px_width / 2
        ColumnCreator.img_line_3.top = pt_draw(1) - px_z_max
        ColumnCreator.img_line_3.Width = px_width
        ColumnCreator.img_line_3.Height = 2
        ColumnCreator.img_line_3.Visible = True
        
        'Draw plywood seams
        Dim cumulative_z As Double
        For i = 1 To UBound(ply_seams) - 1
            Coll(i).Left = pt_draw(0) - px_width / 2
            cumulative_z = 0
            For j = 1 To i
                cumulative_z = cumulative_z + ply_seams(j)
            Next j
            Coll(i).top = pt_draw(1) - px_z_max * (cumulative_z / z)
            Coll(i).Width = px_width
            Coll(i).Height = 2
            Coll(i).Visible = True
        Next i
        
        'Dimension plywood seams
        For i = 1 To UBound(ply_seams)
            Coll(i + 10).Left = pt_draw(0) + px_width / 2 + 10
            cumulative_z = 0
            For j = 1 To i
                cumulative_z = cumulative_z + ply_seams(j)
            Next j
            Coll(i + 10).top = pt_draw(1) - px_z_max * ((cumulative_z - ply_seams(i) / 2) / z) - Coll(i + 10).Height / 2 + 2
            Coll(i + 10).Caption = ConvertFtIn(ply_seams(i))
            Coll(i + 10).Visible = True
        Next i
        
        'Create window seam
        If ColumnCreator.chkWinX.Value = True Or ColumnCreator.chkWinY.Value = True Then
            'Create values for window positions if first time
            If ColumnCreator.WinDim1.Value = "Z1" Then
                ColumnCreator.WinDim1.Value = ConvertFtIn(Round(z / 5, 0)) 'Just assume window is top 1/4 of column for starters
                ColumnCreator.WinDim2.Value = ConvertFtIn(z - ConvertToNum(ColumnCreator.WinDim1.Value))
            End If
            
            'Catch pour window dimensions that go off screen, change to max height
            If InputSource = 2 Then
                If ConvertToNum(ColumnCreator.WinDim1.Value) > z Then
                    ColumnCreator.WinDim1.Value = ConvertFtIn(z)
                    ColumnCreator.WinDim2.Value = ConvertFtIn(0)
                ElseIf ConvertToNum(ColumnCreator.WinDim2.Value) > z Then
                    ColumnCreator.WinDim2.Value = ConvertFtIn(z)
                    ColumnCreator.WinDim1.Value = ConvertFtIn(0)
                End If
            End If
            
            If InputSource = 0 Or InputSource = 2 Or InputSource = 3 Then
                'Set location of window seam
                ColumnCreator.img_line_win.Left = pt_draw(0) - px_width / 2
                ColumnCreator.img_line_win.top = pt_draw(1) - px_z_max + ConvertToNum(ColumnCreator.WinDim1.Value) * px_z_max / z
                ColumnCreator.img_line_win.Width = px_width
                ColumnCreator.img_line_win.Height = 3 'Thicc
                
                'Move dimensions into place (upper dim)
                ColumnCreator.WinDim1.Left = pt_draw(0) - px_width / 2 - 40
                ColumnCreator.WinDim1.top = pt_draw(1) - px_z_max + (ConvertToNum(ColumnCreator.WinDim1.Value) * px_z_max / z) / 2 - ColumnCreator.WinDim1.Height / 2
                'Move dimensions into place (lower dim)
                ColumnCreator.WinDim2.Left = pt_draw(0) - px_width / 2 - 40
                ColumnCreator.WinDim2.top = pt_draw(1) - (ConvertToNum(ColumnCreator.WinDim2.Value) * px_z_max / z) / 2 - ColumnCreator.WinDim1.Height / 2
                
                'If the window positioning dimensions add up to the height, make the dimensions blue, otherwise, grey
UpdatePlyColorCheck:
                If Round(ConvertToNum(ColumnCreator.WinDim1.Value) + ConvertToNum(ColumnCreator.WinDim2.Value), 5) = Round(z, 5) Then
                    ColumnCreator.WinDim1.ForeColor = &HFF0000
                    ColumnCreator.WinDim2.ForeColor = &HFF0000
                Else
                    ColumnCreator.WinDim1.ForeColor = &H80000011
                    ColumnCreator.WinDim2.ForeColor = &H80000011
                End If
            
            End If

            'Make window seam and dimensions visible
            ColumnCreator.img_line_win.Visible = True
            ColumnCreator.WinDim1.Visible = True
            ColumnCreator.WinDim2.Visible = True
        Else
            'Hide window dimensions if no window option is checked
            ColumnCreator.WinDim1.Visible = False
            ColumnCreator.WinDim2.Visible = False
        End If
    Else
        'Hide window dimensions if missing inputs
        ColumnCreator.WinDim1.Visible = False
        ColumnCreator.WinDim2.Visible = False
    End If
    
UpdatePlyLastLine:
End Function

'===============================================================================================================================
'   Other useful functions
'===============================================================================================================================

Public Function DisableDrawingButtons()
    'Disable drawing buttons
    ColumnCreator.CommandButton1.Enabled = False
    ColumnCreator.CommandButton1.BackColor = &H80000004 'grey
    ColumnCreator.CommandButton1.ForeColor = &H80000000 'background grey
    ColumnCreator.CommandButton1.Enabled = False
    ColumnCreator.CommandButton1.BackColor = &H80000004 'grey
    ColumnCreator.CommandButton1.ForeColor = &H80000000 'background grey
End Function
Public Function EnableDrawingButtons()
    'Enable drawing buttons
    ColumnCreator.CommandButton1.Enabled = True
    ColumnCreator.CommandButton1.BackColor = &H8000000D 'blue
    ColumnCreator.CommandButton1.ForeColor = &H80000012 'black
    ColumnCreator.CommandButton1.Enabled = True
    ColumnCreator.CommandButton1.BackColor = &H8000000D 'blue
    ColumnCreator.CommandButton1.ForeColor = &H80000012 'black
End Function
Public Function ValidatePlySeams(ply_seams() As Double, x As Double, y As Double, z As Double) As Integer
'Returns 0: not valid
'Returns 1: valid
    
    'Assume layout is valid
    ValidatePlySeams = 1
    
    'Check whether the sum of listed lengths equals the column height
    Dim temp_ht As Double: temp_ht = 0
    For i = 1 To UBound(ply_seams)
        temp_ht = temp_ht + ply_seams(i)
    Next i
    If Round(temp_ht, 3) <> Round(z, 3) Then
        If temp_ht > z Then
            ColumnCreator.txtPlyError.Caption = "Remove " & ConvertFtIn(temp_ht - z)
        ElseIf temp_ht < z Then
            ColumnCreator.txtPlyError.Caption = "Add " & ConvertFtIn(z - temp_ht)
        End If
        ColumnCreator.txtPlyError.Visible = True
        ValidatePlySeams = 0
        Call DisableDrawingButtons
        GoTo ValidatePlySeamsLastLine
    End If
    
    'Check that plywood doesn't exceed max size
    Dim ply_width_x As Double
    Dim ply_width_y As Double
    Dim min_ply_ht As Double: min_ply_ht = 6
    Dim max_ply_ht As Double
    
    ply_width_x = x + 1.5
    ply_width_y = y + 1.5
    If Round(ply_width_x, 3) > 48 Or Round(ply_width_y, 3) > 48 Then
        max_ply_ht = 48
    Else
        max_ply_ht = 96
    End If
    
    For i = 1 To UBound(ply_seams)
        If ply_seams(i) > max_ply_ht Then
            ColumnCreator.txtPlyError.Caption = ConvertFtIn(ply_seams(i)) & " ply too tall. Max: " & ConvertFtIn(max_ply_ht)
            ColumnCreator.txtPlyError.Visible = True
            ValidatePlySeams = 0
            Call DisableDrawingButtons
            GoTo ValidatePlySeamsLastLine
        ElseIf ply_seams(i) < min_ply_ht Then
            ColumnCreator.txtPlyError.Caption = ConvertFtIn(ply_seams(i)) & " ply too small. Min: " & ConvertFtIn(min_ply_ht)
            ColumnCreator.txtPlyError.Visible = True
            ValidatePlySeams = 0
            Call DisableDrawingButtons
            GoTo ValidatePlySeamsLastLine
        Else
            ColumnCreator.txtPlyError.Visible = False
        End If
    Next i

'Hide error text if no errors have been found here
If ValidatePlySeams = 1 Then
    ColumnCreator.txtPlyError.Visible = False
End If

ValidatePlySeamsLastLine:
End Function

    }
}*/