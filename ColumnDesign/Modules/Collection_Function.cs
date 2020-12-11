/*namespace ColumnDesign.Modules
{
    public class Collection_Function
    {
        Public Function BuildCollection() As Collection
    'Create collection of image objects for drawing the elevation view
    
    Dim Coll As Collection
    Set Coll = New Collection
    
    If ColumnCreator.MultiPage1.Value = 0 Then
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'                    GATES COLUMNS
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
        Coll.Add ColumnCreator.img_line_4   'Item #1
        Coll.Add ColumnCreator.img_line_5
        Coll.Add ColumnCreator.img_line_6
        Coll.Add ColumnCreator.img_line_7
        Coll.Add ColumnCreator.img_line_8
        Coll.Add ColumnCreator.img_line_9
        Coll.Add ColumnCreator.img_line_10
        Coll.Add ColumnCreator.img_line_11
        Coll.Add ColumnCreator.img_line_12
        Coll.Add ColumnCreator.img_line_13
        Coll.Add ColumnCreator.txt_1        'Item #11
        Coll.Add ColumnCreator.txt_2
        Coll.Add ColumnCreator.txt_3
        Coll.Add ColumnCreator.txt_4
        Coll.Add ColumnCreator.txt_5
        Coll.Add ColumnCreator.txt_6
        Coll.Add ColumnCreator.txt_7
        Coll.Add ColumnCreator.txt_8
        Coll.Add ColumnCreator.txt_9
        Coll.Add ColumnCreator.txt_10
        Coll.Add ColumnCreator.txt_11
    ElseIf ColumnCreator.MultiPage1.Value = 1 Then
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
'                 SCISSOR CLAMP COLUMNS
' # # # # # # # # # # # # # # # # # # # # # # # # # # # #
        Coll.Add ColumnCreator.s_img_line_4   'Item #1
        Coll.Add ColumnCreator.s_img_line_5
        Coll.Add ColumnCreator.s_img_line_6
        Coll.Add ColumnCreator.s_img_line_7
        Coll.Add ColumnCreator.s_img_line_8
        Coll.Add ColumnCreator.s_img_line_9
        Coll.Add ColumnCreator.s_img_line_10
        Coll.Add ColumnCreator.s_img_line_11
        Coll.Add ColumnCreator.s_img_line_12
        Coll.Add ColumnCreator.s_img_line_13
        Coll.Add ColumnCreator.s_txt_1        'Item #11
        Coll.Add ColumnCreator.s_txt_2
        Coll.Add ColumnCreator.s_txt_3
        Coll.Add ColumnCreator.s_txt_4
        Coll.Add ColumnCreator.s_txt_5
        Coll.Add ColumnCreator.s_txt_6
        Coll.Add ColumnCreator.s_txt_7
        Coll.Add ColumnCreator.s_txt_8
        Coll.Add ColumnCreator.s_txt_9
        Coll.Add ColumnCreator.s_txt_10
        Coll.Add ColumnCreator.s_txt_11
    End If
    
    Set BuildCollection = Coll
    
End Function

    }
}*/