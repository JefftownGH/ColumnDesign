using ColumnDesign.UI;
using ColumnDesign.ViewModel;

namespace ColumnDesign.Modules
{
    public class sUpdatePly_Function
    {
        public static void sUpdatePly(ColumnCreatorView ui, ColumnCreatorViewModel vm)
        {
        }

//         Public Function sUpdatePly(InputSource As Integer)
//
//     Dim x As Double
//     Dim y As Double
//     Dim z As Double
//     
//     x = ConvertToNum(ColumnCreator.sWidthBox)
//     y = ConvertToNum(ColumnCreator.sLengthBox)
//     z = ConvertToNum(ColumnCreator.sHeightBox)

//     ColumnCreator.s_img_line_1.Visible = False
//     ColumnCreator.s_img_line_2.Visible = False
//     ColumnCreator.s_img_line_3.Visible = False
//     For i = 1 To Coll.Count
//         Coll.Item(i).Visible = False
//     Next i
//     
//     'Triggers only when x, y, and z are all provided
//     If Round(x * y * z, 3) <> 0 Then
//     
//         'Check that x, y, and z inputs are valid. If not, skip the rest of this function. (qtyCheck = 0 means the quantity box is not considered)
//         If CheckInputs(ConvertToNum(ColumnCreator.sWidthBox.Value), ConvertToNum(ColumnCreator.sLengthBox.Value), ConvertToNum(ColumnCreator.sHeightBox.Value), ConvertToNum(ColumnCreator.sQuantityBox.Value), 0) = 0 Then
//             Call sDisableDrawingButtons
//             GoTo UpdatePlyLastLine
//         End If
//     
//         'Enable drawing buttons (may be disabled later if ply seams are invalid)
//         Call sEnableDrawingButtons
//         
// '===============================================================================================================================
// '   Manage userform objects and ply_seams textbox
// '===============================================================================================================================
//
//         'Get plywood seams layout
//         Dim ply_seams() As Double
//         Dim temp_ht As Double: temp_ht = 0
//
//         If InputSource = 3 Or ColumnCreator.sboxPlySeams.Value = "" Then
//             'Generate the ply seams
//             ply_seams = GetPlySeams(x, y, z)
//             
//             'Hide any previous error messages because these fresh ply seams should be flawless
//             ColumnCreator.stxtPlyError.Visible = False
//         ElseIf InputSource = 0 Or InputSource = 1 Then
//             'Read the ply seams
//             ply_seams = ReadSizes(ColumnCreator.sboxPlySeams.Value)
//             
//             'If ply seams are invalid: disable the drawing buttons and skip the rest of this function
//             If sValidatePlySeams(ply_seams, x, y, z) <> 1 Then
//                 Call sDisableDrawingButtons
//                 GoTo UpdatePlyLastLine
//             End If
//         End If
//         
//         'If quantity is missing, disable drawing buttons and continue
//         If ConvertToNum(ColumnCreator.sQuantityBox.Value) = 0 Then
//             Call sDisableDrawingButtons
//         End If
//             
//         'Populate the textbox with a new plywood layout
//         If InputSource = 0 Or InputSource = 3 Then
//             'If input is from the x or y box and there's already a ply layout, don't replace it
//             If InputSource = 0 And ColumnCreator.sboxPlySeams.Value <> "" Then
//                 GoTo SkipPlyUpdate
//             End If
//             
//             Dim strPlySeams As String
//             For i = 1 To UBound(ply_seams)
//                 If i <> UBound(ply_seams) Then
//                     strPlySeams = strPlySeams & ConvertFtIn(ply_seams(i)) & ", "
//                 ElseIf i = UBound(ply_seams) Then
//                     strPlySeams = strPlySeams & ConvertFtIn(ply_seams(i))
//                 End If
//             Next i
//             ColumnCreator.sboxPlySeams.Value = strPlySeams
//         End If
// SkipPlyUpdate:
//
// '===============================================================================================================================
// '   Draw the column on the userform
// '===============================================================================================================================
//         'Define origin point for column sketch
//         Dim pt_draw(0 To 2) As Double
//         pt_draw(0) = 200
//         pt_draw(1) = 216
//         pt_draw(2) = 0
//         
//         'Define max pixel height and actual pixel width for x and y sides
//         Dim px_z_max As Double
//         Dim px_x As Double
//         Dim px_y As Double
//         If z > 96 Then 'Display column at full height and modify the width to scale
//             px_z_max = 200
//             px_x = px_z_max * (x / z)
//             px_y = px_z_max * (y / z)
//         Else 'Scale down display height and width to fit. Without this step, a short column will be displayed as extremely wide and go off screen.
//             px_z_max = 200 * z / 96
//             px_x = px_z_max * (x / z) * z / 96
//             px_y = px_z_max * (y / z) * z / 96
//         End If
//         
//         'Change how this is defined for view types later
//         If ColumnCreator.slblAxis.Caption = "X" Then
//             px_width = px_x
//         ElseIf ColumnCreator.slblAxis.Caption = "Y" Then
//             px_width = px_y
//         Else
//             MsgBox ("Error: Horizontal dimension not x or y")
//         End If
//         
//         'Draw "frame" of column: left, right, and top
//         'Left
//         ColumnCreator.s_img_line_1.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.s_img_line_1.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_1.Width = 2
//         ColumnCreator.s_img_line_1.Height = px_z_max
//         ColumnCreator.s_img_line_1.Visible = True
//     
//         'Right
//         ColumnCreator.s_img_line_2.Left = pt_draw(0) + px_width / 2
//         ColumnCreator.s_img_line_2.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_2.Width = 2
//         ColumnCreator.s_img_line_2.Height = px_z_max
//         ColumnCreator.s_img_line_2.Visible = True
//     
//         'Top
//         ColumnCreator.s_img_line_3.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.s_img_line_3.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_3.Width = px_width
//         ColumnCreator.s_img_line_3.Height = 2
//         ColumnCreator.s_img_line_3.Visible = True
//         
//         'Draw plywood seams
//         Dim cumulative_z As Double
//         For i = 1 To UBound(ply_seams) - 1
//             Coll(i).Left = pt_draw(0) - px_width / 2
//             cumulative_z = 0
//             For j = 1 To i
//                 cumulative_z = cumulative_z + ply_seams(j)
//             Next j
//             Coll(i).top = pt_draw(1) - px_z_max * (cumulative_z / z)
//             Coll(i).Width = px_width
//             Coll(i).Height = 2
//             Coll(i).Visible = True
//         Next i
//         
//         'Dimension plywood seams
//         For i = 1 To UBound(ply_seams)
//             Coll(i + 10).Left = pt_draw(0) + px_width / 2 + 10
//             cumulative_z = 0
//             For j = 1 To i
//                 cumulative_z = cumulative_z + ply_seams(j)
//             Next j
//             Coll(i + 10).top = pt_draw(1) - px_z_max * ((cumulative_z - ply_seams(i) / 2) / z) - Coll(i + 10).Height / 2 + 2
//             Coll(i + 10).Caption = ConvertFtIn(ply_seams(i))
//             Coll(i + 10).Visible = True
//         Next i
//     End If
//     
// UpdatePlyLastLine:
// End Function
//
// '===============================================================================================================================
// '   Other useful functions
// '===============================================================================================================================
//
        public static int sValidatePlySeams(ColumnCreatorView ui, double[] ply_seams, double x, double y,
            double z)
        {
            
// 'Returns 0: not valid
// 'Returns 1: valid
//     
//     'Assume layout is valid
//     sValidatePlySeams = 1
//     
//     'Check whether the sum of listed lengths equals the column height
//     Dim temp_ht As Double: temp_ht = 0
//     For i = 1 To UBound(ply_seams)
//         temp_ht = temp_ht + ply_seams(i)
//     Next i
//     If Round(temp_ht, 3) <> Round(z, 3) Then
//         If temp_ht > z Then
//             ColumnCreator.stxtPlyError.Caption = "Remove " & ConvertFtIn(temp_ht - z)
//         ElseIf temp_ht < z Then
//             ColumnCreator.stxtPlyError.Caption = "Add " & ConvertFtIn(z - temp_ht)
//         End If
//         ColumnCreator.stxtPlyError.Visible = True
//         sValidatePlySeams = 0
//         GoTo sValidatePlySeamsLastLine
//     End If
//     
//     'Check that plywood doesn't exceed max size
//     Dim ply_width_x As Double
//     Dim ply_width_y As Double
//     Dim min_ply_ht As Double: min_ply_ht = 6
//     Dim max_ply_ht As Double
//     
//     ply_width_x = x
//     ply_width_y = y + 4.5
//     If Round(ply_width_x, 3) > 48 Or Round(ply_width_y, 3) > 48 Then
//         max_ply_ht = 48
//     Else
//         max_ply_ht = 96
//     End If
//     
//     For i = 1 To UBound(ply_seams)
//         If ply_seams(i) > max_ply_ht Then
//             ColumnCreator.stxtPlyError.Caption = ConvertFtIn(ply_seams(i)) & " ply too tall. Max: " & ConvertFtIn(max_ply_ht)
//             ColumnCreator.stxtPlyError.Visible = True
//             sValidatePlySeams = 0
//             GoTo sValidatePlySeamsLastLine
//         ElseIf ply_seams(i) < min_ply_ht Then
//             ColumnCreator.stxtPlyError.Caption = ConvertFtIn(ply_seams(i)) & " ply too small. Min: " & ConvertFtIn(min_ply_ht)
//             ColumnCreator.stxtPlyError.Visible = True
//             sValidatePlySeams = 0
//             GoTo sValidatePlySeamsLastLine
//         Else
//             ColumnCreator.stxtPlyError.Visible = False
//         End If
//     Next i
//
// 'Hide error text if no errors have been found here
// If sValidatePlySeams = 1 Then
//     ColumnCreator.stxtPlyError.Visible = False
// End If
//
// sValidatePlySeamsLastLine:
            return 0;
        }
//
// Public Function sCheckHeight(x As Double, y As Double, z As Double)
//     
//     Dim big_dim As Double
//     Dim max_ht As Double
//     
//     'Check that x, and y are provided. If either are missing, hide the max height text and don't do this function
//     'The max height text is also hidden from the CheckInputs_Function
//     If x = 0 Or y = 0 Then
//         ColumnCreator.slblMaxHt.Caption = "N/A"
//         Exit Function
//     End If
//     
//     'Find the longer side of the column
//     If x >= y Then
//         big_dim = x
//     Else
//         big_dim = y
//     End If
//     
//     'Use a pre-built table to set the maximum height
//     Select Case Round(big_dim, 3)
//     Case Is <= 8
//         max_ht = 196
//     Case Is <= 10
//         max_ht = 196
//     Case Is <= 12
//         max_ht = 196
//     Case Is <= 14
//         max_ht = 196
//     Case Is <= 16
//         max_ht = 196
//     Case Is <= 18
//         max_ht = 197
//     Case Is <= 20
//         max_ht = 197
//     Case Is <= 22
//         max_ht = 193
//     Case Is <= 24
//         max_ht = 195
//     Case Is <= 26
//         max_ht = 175
//     Case Is <= 28
//         max_ht = 175
//     Case Is <= 30
//         max_ht = 175
//     Case Is <= 32
//         max_ht = 133
//     Case Is <= 34
//         max_ht = 133
//     Case Is <= 36
//         max_ht = 180
//     Case Is <= 38
//         max_ht = 127
//     Case Is <= 40
//         max_ht = 127
//     Case Is <= 42
//         max_ht = 127
//     Case Is <= 44
//         max_ht = 127
//     Case Is <= 46
//         max_ht = 127
//     End Select
//
//     'Display the max height on the userform
//     ColumnCreator.slblMaxHt.Caption = ConvertFtIn(max_ht)
//     
//     'Run check inputs
//     Call CheckInputs(ConvertToNum(ColumnCreator.sWidthBox.Value), ConvertToNum(ColumnCreator.sLengthBox.Value), ConvertToNum(ColumnCreator.sHeightBox.Value), ConvertToNum(ColumnCreator.sQuantityBox.Value), 1)
//     
// End Function
//
     }
}