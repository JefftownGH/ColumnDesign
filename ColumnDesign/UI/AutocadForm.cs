/*

Private Sub CommandButton2_Click()
    'Call scissor clamp module
    Call CreateScissorClamp
End Sub

Private Sub CommandButton1_Click()
        
    '#####################################################################################
    '                                 D R A W I N G
    '#####################################################################################

    Dim hatchObj As AcadHatch
    Dim OuterLoop(0) As AcadEntity
    Dim pl_pts(0 To 9) As Double
    
    'Set boundry of hatch (must close hatch by returning to initial point)
    pl_pts(0) = pt13(0):     pl_pts(1) = pt13(1)
    pl_pts(2) = pt13(0) - x: pl_pts(3) = pt13(1)
    pl_pts(4) = pt13(0) - x: pl_pts(5) = pt13(1) + y
    pl_pts(6) = pt13(0):     pl_pts(7) = pt13(1) + y
    pl_pts(8) = pt13(0):     pl_pts(9) = pt13(1)
    
    Set OuterLoop(0) = ThisDrawing.ModelSpace.AddLightWeightPolyline(pl_pts) 'Assign the pline object as the hatch's outer loop (inner loop optional and skipped here)
    Set hatchObj = ThisDrawing.ModelSpace.AddHatch(0, "AR-CONC", True) 'Set hatch
    hatchObj.AppendOuterLoop (OuterLoop)
    OuterLoop(0).Delete
    
    pl_pts(1) = pl_pts(1) + 83.5: pl_pts(3) = pl_pts(3) + 83.5: pl_pts(5) = pl_pts(5) + 83.5: pl_pts(7) = pl_pts(7) + 83.5: pl_pts(9) = pl_pts(9) + 83.5 'Move pts to top clamps' plan
    Set OuterLoop(0) = ThisDrawing.ModelSpace.AddLightWeightPolyline(pl_pts)
    Set hatchObj = ThisDrawing.ModelSpace.AddHatch(0, "AR-CONC", True)
    hatchObj.AppendOuterLoop (OuterLoop)
    OuterLoop(0).Delete
    
    'Draw text in plan views
    ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
    
    'Shrink text for small columns
    If x <= 16 Or y <= 16 Then
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 0.75 + y / 2
        Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt17, x, "CLAMPS")
        MTextObject2.StyleName = "Arial Narrow Bold";
        MTextObject2.Height = 1.5: MTextObject2.AttachmentPoint = acAttachmentPointTopCenter
    
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 2.1 + y / 2 + 83.5
        Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt17, x, "TOP" & vbNewLine & "CLAMPS")
        MTextObject2.StyleName = "Arial Narrow Bold": MTextObject2.Height = 1.5: MTextObject2.AttachmentPoint = acAttachmentPointTopCenter
    Else
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 1.25 + y / 2
        Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt17, x, "CLAMPS")
        MTextObject2.StyleName = "Arial Narrow Bold": MTextObject2.Height = 2.5: MTextObject2.AttachmentPoint = acAttachmentPointTopCenter
    
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 3.35 + y / 2 + 83.5
        Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt17, x, "TOP" & vbNewLine & "CLAMPS")
        MTextObject2.StyleName = "Arial Narrow Bold": MTextObject2.Height = 2.5: MTextObject2.AttachmentPoint = acAttachmentPointTopCenter
    End If
        
    For i = 0 To 1
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + y - 0.5 + 83.5 * i
        Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, "SIDE A")
        MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75
        
        If DrawB = True Then
            pt17(0) = pt13(0) - x + 0.5: pt17(1) = pt13(1) + 83.5 * i
            Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, "SIDE B")
            MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75
        End If
        
        pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 2.25 + 83.5 * i
        Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, x)
        MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75
        
        pt17(0) = pt13(0) - 2.25: pt17(1) = pt13(1) + 83.5 * i
        Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, y)
        MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75
    Next i
    
    'Add all static lines, text, objects, and images as one block
    'For tall columns use the "high" background to accomdate the column size
    'For smaller columns use the "medium" background to avoid the unaesthetic blank space between a relatively short column and the top of page details
    If z > 190 Then
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_o, "VBA_COLUMN_BACKGROUND_HIGH", 1#, 1#, 1#, 0)
    Else
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_o, "VBA_COLUMN_BACKGROUND_MEDIUM", 1#, 1#, 1#, 0)
    End If
    
    'Add alternative picking eye deatils if no window is used
    If window = False Then
        pt_blk(0) = pt_o(0) + 402: pt_blk(1) = pt_o(1) + 261 - 36 * Abs(z <= 190)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_COLUMN_ALT_PICKING_DETAILS", 1#, 1#, 1#, 0)
    End If
    
    'Add detail references to each elevation view
    ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
    
    'Side A detail reference
    pt_blk(0) = ptA(0) + ply_width_x / 2: pt_blk(1) = pt_o(1) + 9
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
    AttList = BlockRefObj.GetAttributes
    For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
        If AttList(i).TextString = "X" Then
            AttList(i).TextString = "A"
        End If
        If AttList(i).TextString = "DESCRIPTION" Then
            AttList(i).TextString = "SIDE ""A"" PANEL"
        End If
        If AttList(i).TextString = "REFERENCE" Then
            AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
        End If
    Next i
    
    'Side B detail reference
    If DrawB = True Then
        pt_blk(0) = ptB(0) + ply_width_y / 2: pt_blk(1) = pt_o(1) + 9
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
        AttList = BlockRefObj.GetAttributes
        For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
            If AttList(i).TextString = "X" Then
                AttList(i).TextString = "B"
            End If
            If AttList(i).TextString = "DESCRIPTION" Then
                AttList(i).TextString = "SIDE ""B"" PANEL"
            End If
            If AttList(i).TextString = "REFERENCE" Then
                AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
            End If
        Next i
    End If

    'Window detail reference
    If DrawW = True Then
        pt_blk(0) = ptW(0) + ply_width_w / 2: pt_blk(1) = pt_o(1) + 9
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
        AttList = BlockRefObj.GetAttributes
        For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
            If AttList(i).TextString = "X" Then
                AttList(i).TextString = "W"
            End If
            If AttList(i).TextString = "DESCRIPTION" Then
                AttList(i).TextString = "POUR WINDOW PANEL"
            End If
            If AttList(i).TextString = "REFERENCE" Then
                AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
            End If
        Next i
    End If
    
    'Elevation view detail reference
    pt_blk(0) = ptE(0) + ply_width_e / 2: pt_blk(1) = pt_o(1) + 9
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
    AttList = BlockRefObj.GetAttributes
    For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
        If AttList(i).TextString = "X" Then
            AttList(i).TextString = "E"
        End If
        If AttList(i).TextString = "DESCRIPTION" Then
            AttList(i).TextString = "COLUMN FORM ELEVATION"
        End If
        If AttList(i).TextString = "REFERENCE" Then
            AttList(i).TextString = ""
        End If
    Next i
    
    'Add lower left fab count below "column form elevation"
    ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
    count_str = "FAB " & n_col & "-EA"
    pt32(0) = ptE(0) + ply_width_e / 2: pt32(1) = pt_o(1) + 7
    Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt32, 30, count_str)
    MTextObject2.Height = 2.5
    
    
'#####################################################################################
'                                S H E E T I N G
'#####################################################################################

    'Insert border
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_click, "VBA_24X36_LAYOUT", 1#, 1#, 1#, 0) 'Insert 3/4" = 1'-0" sheet boundry
    BlockRefObj.ScaleEntity pt_click, 12 / 0.75
    AttList = BlockRefObj.GetAttributes
    For i = LBound(AttList) To UBound(AttList)
        If AttList(i).TextString = "SCALE" Then
            AttList(i).TextString = "3/4'' = 1'-0''"
        
        'Use a custom scale
        objPSViewPort.StandardScale = acVpCustomScale
        objPSViewPort.CustomScale = 0.04136364
        

'#####################################################################################
'                   M I S C E L L A N E O U S  &  C L E A N  U P
'#####################################################################################

    'If custom clamp size is selected but too small, warn the user
    If ClampSizeButton1.Value = True Then
        If x > 24 Or y > 24 Then
            MsgBox ("Warning: Selected clamp appears too small for this column")
        End If
    ElseIf ClampSizeButton2.Value = True Then
        If x > 36 Or y > 36 Then
            MsgBox ("Warning: Selected clamp appears too small for this column")
        End If
    ElseIf ClampSizeButton3.Value = True Then
        If x > 48 Or y > 48 Then
            MsgBox ("Warning: Selected clamp appears too small for this column")
        End If
    End If
    End sub


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
Public Function DrawDim(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double)

    Dim dimObj As AcadDimAligned
    Dim point1(0 To 2) As Double 'first point
    Dim point2(0 To 2) As Double 'second point
    Dim location(0 To 2) As Double 'text location
    
    ' Define the dimension
    point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
    point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
    location(0) = x3#:  location(1) = y3#:  location(2) = 0#
    
    ' Create an aligned dimension object in model space
    Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
End Function
Public Function DrawDimSuffix(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, txt As String)

    Dim dimObj As AcadDimAligned
    Dim point1(0 To 2) As Double
    Dim point2(0 To 2) As Double
    Dim location(0 To 2) As Double
    
    ' Define the dimension
    point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
    point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
    location(0) = x3#:  location(1) = y3#:  location(2) = 0#
    
    ' Create an aligned dimension object in model space
    Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
    dimObj.TextSuffix = txt
    
End Function
Public Function DrawDimLin(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, angle As Double)

    Dim dimObj As AcadDimRotated
    Dim point1(0 To 2) As Double 'first point
    Dim point2(0 To 2) As Double 'second point
    Dim location(0 To 2) As Double 'text location
    
    ' Define the dimension
    point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
    point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
    location(0) = x3#:  location(1) = y3#:  location(2) = 0#
    
    ' Create an aligned dimension object in model space
    Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
End Function
Public Function DrawDimLinSuffixLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, txt As String, angle As Double)

    Dim dimObj As AcadDimRotated
    Dim point1(0 To 2) As Double
    Dim point2(0 To 2) As Double
    Dim location(0 To 2) As Double
    Dim txtLocation(0 To 2) As Double
    
    ' Define the dimension
    point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
    point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
    location(0) = x3#: location(1) = y3#: location(2) = 0#
    txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
    
    ' Create an aligned dimension object in model space
    Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
    dimObj.TextMovement = acMoveTextAddLeader
    dimObj.TextPosition = txtLocation
    dimObj.TextSuffix = txt
End Function
Public Function DrawDimLinLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, angle As Double)

    Dim dimObj As AcadDimRotated
    Dim point1(0 To 2) As Double
    Dim point2(0 To 2) As Double
    Dim location(0 To 2) As Double
    Dim txtLocation(0 To 2) As Double
    
    ' Define the dimension
    point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
    point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
    location(0) = x3#: location(1) = y3#: location(2) = 0#
    txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
    
    ' Create an aligned dimension object in model space
    Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
    dimObj.TextMovement = acMoveTextAddLeader
    dimObj.TextPosition = txtLocation
End Function

Private Sub SaveButton_Click()
    Dim SheetDataFilePath As String
    Dim SheetDataFileName As String
    Dim SheetDataFilePathAndName As String
    Dim sHostName As String
    Dim fso As New FileSystemObject
    Dim oFile As Object
    
    'Set file name and path
    SheetDataFilePath = "C:\ProgramData\Autodesk\VBA Saved Files"
    SheetDataFileName = "Column Sheet Data.txt"
    SheetDataPathAndName = SheetDataFilePath & "\" & SheetDataFileName
    
    'Create folder path
    If Not FolderExists("C:\ProgramData") Then
        FolderCreate ("C:\ProgramData")
    End If
    If Not FolderExists("C:\ProgramData\Autodesk") Then
        FolderCreate ("C:\ProgramData\Autodesk")
    End If
    If Not FolderExists(SheetDataFilePath) Then
        FolderCreate (SheetDataFilePath)
    End If

    'Create a text file and write the following lines
    Set oFile = fso.CreateTextFile(SheetDataPathAndName, True) 'True allows overwriting
    
    oFile.WriteLine ProjectNameBox.Value
    oFile.WriteLine ProjectTitleBox.Value
    oFile.WriteLine ProjectAddressBox.Value
    oFile.WriteLine DateBox.Value
    oFile.WriteLine SheetIssuedForBox.Value
    oFile.WriteLine ScaleBox.Value
    oFile.WriteLine DrawnByBox.Value
    oFile.WriteLine JobNoBox.Value
    oFile.WriteLine SheetBox.Value
    oFile.WriteLine SuffixBox.Value
    oFile.WriteLine AreaBox.Value
    
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub

Private Sub LoadButton_Click()
    Dim SheetDataFilePath As String
    Dim SheetDataFileName As String
    Dim SheetDataFilePathAndName As String
    Dim mystring As String
    Dim temp_lines(1 To 11) As String
    
    SheetDataFilePath = "C:\ProgramData\Autodesk\VBA Saved Files"
    SheetDataFileName = "Column Sheet Data.txt"
    SheetDataPathAndName = SheetDataFilePath & "\" & SheetDataFileName
    
    If Dir(SheetDataPathAndName) = "" Then
        MsgBox ("Error: You can't load a file you haven't saved")
        Exit Sub
    End If
    
    Open SheetDataPathAndName For Input As #1
    i = 1
    
    Do While Not EOF(1)
        Line Input #1, mystring
        temp_lines(i) = mystring
        i = i + 1
    Loop
    Close #1
    
    'Load the following lines to the corresponding variables
    ProjectNameBox.Value = temp_lines(1)
    ProjectTitleBox.Value = temp_lines(2)
    ProjectAddressBox.Value = temp_lines(3)
    DateBox.Value = temp_lines(4)
    SheetIssuedForBox.Value = temp_lines(5)
    ScaleBox.Value = temp_lines(6)
    DrawnByBox.Value = temp_lines(7)
    JobNoBox.Value = temp_lines(8)
    SheetBox.Value = temp_lines(9)
    SuffixBox.Value = temp_lines(10)
    AreaBox.Value = temp_lines(11)
End Sub

DeadInTheWater:
    MsgBox "A folder could not be created for the following path: " & path & ". Check the path name and try again."
    FolderCreate = False
    Exit Function
End Function

Private Sub UserForm_Initialize()


'Show a warning when plywood is changed from HDO
Private Sub PlyNameBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub PlyNameBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    If UCase(ColumnCreator.PlyNameBox.Value) <> "HDO" Then
        ColumnCreator.notHDOwarn.Visible = True
    Else
        ColumnCreator.notHDOwarn.Visible = False
    End If
End Sub

Private Sub ProjectNameBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If LCase(ProjectNameBox.Value) = "open" Then
        MultiPage1.Pages.Item("Page2").Visible = True
    End If
    If LCase(ProjectNameBox.Value) = "close" Then
        MultiPage1.Pages.Item("Page2").Visible = False
    End If
End Sub


'[-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-]
'                        Scissor clamp section
'[-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-][-=-]


Private Sub sWidthBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Update the plywood layout sketch
    Call sUpdatePly.sUpdatePly(0)

    'Check that x, y, z inputs are within the defined limits
    If CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 0 Then
        Call sUpdatePly.sDisableDrawingButtons
    ElseIf CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 1 Then
        Call sUpdatePly.sEnableDrawingButtons
    End If
    
    'Update max allowable height
    Call sUpdatePly.sCheckHeight(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value))
    
    'If an error message is visible, disable the drawing button
    If ColumnCreator.stxtPlyError.Visible = True Then
        Call sUpdatePly.sDisableDrawingButtons
    End If
    
    'Auto-generate sheet name
    Call AutoSheetName
    
End Sub
Private Sub sLengthBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Update the plywood layout sketch
    Call sUpdatePly.sUpdatePly(0)

    'Check that x, y, z inputs are within the defined limits
    If CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 0 Then
        Call sUpdatePly.sDisableDrawingButtons
    ElseIf CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 1 Then
        Call sUpdatePly.sEnableDrawingButtons
    End If
    
    'Update max allowable height
    Call sUpdatePly.sCheckHeight(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value))

    'If an error message is visible, disable the drawing button
    If ColumnCreator.stxtPlyError.Visible = True Then
        Call sUpdatePly.sDisableDrawingButtons
    End If
    
    'Auto-generate sheet name
    Call AutoSheetName
    
End Sub
Private Sub sHeightBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Update the plywood layout sketch
    Call sUpdatePly.sUpdatePly(0)
    Call sUpdatePly.sUpdatePly(3)

    'Check that x, y, z inputs are within the defined limits
    If CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 0 Then
        Call sUpdatePly.sDisableDrawingButtons
    ElseIf CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 1 Then
        Call sUpdatePly.sEnableDrawingButtons
    End If
    
    'If an error message is visible, disable the drawing button
    If ColumnCreator.stxtPlyError.Visible = True Then
        Call sUpdatePly.sDisableDrawingButtons
    End If
    
    'Update max allowable height
    Call sUpdatePly.sCheckHeight(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value))
    
End Sub
Private Sub sQuantityBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Check that x, y, z inputs are within the defined limits
    If CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 0 Then
        Call sUpdatePly.sDisableDrawingButtons
    ElseIf CheckInputs(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value), ConvertToNum(sQuantityBox.Value), 1) = 1 Then
        Call sUpdatePly.sEnableDrawingButtons
    End If
    
    'If an error message is visible, disable the drawing button
    If ColumnCreator.stxtPlyError.Visible = True Then
        Call sUpdatePly.sDisableDrawingButtons
    End If
    
    'Update max allowable height. Calling this function here to prevent changes the quantity from erasing data shown previously.
    Call sUpdatePly.sCheckHeight(ConvertToNum(sWidthBox.Value), ConvertToNum(sLengthBox.Value), ConvertToNum(sHeightBox.Value))
    
End Sub

Private Sub slblAxis_Click()
    If ColumnCreator.slblAxis.Caption = "X" Then
        ColumnCreator.slblAxis.Caption = "Y"
    ElseIf ColumnCreator.slblAxis.Caption = "Y" Then
        ColumnCreator.slblAxis.Caption = "X"
    End If
    Call sUpdatePly.sUpdatePly(0) 'Update the plywood layout sketch
End Sub

Private Sub sWidthBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
    Call KeyInputRestrictorFtIn(KeyAscii, sWidthBox.Value)
End Sub
Private Sub sLengthBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
    Call KeyInputRestrictorFtIn(KeyAscii, sLengthBox.Value)
End Sub
Private Sub sHeightBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
    Call KeyInputRestrictorFtIn(KeyAscii, sHeightBox.Value)
End Sub
Private Sub sQuantityBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
    Call KeyInputRestrictorNum(KeyAscii)
End Sub

'Show a warning when plywood is changed from HDO
Private Sub sPlyNameBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sPlyNameBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    If UCase(ColumnCreator.sPlyNameBox.Value) <> "HDO" Then
        ColumnCreator.snotHDOwarn.Visible = True
    Else
        ColumnCreator.snotHDOwarn.Visible = False
    End If
End Sub

Private Sub sClampSizeButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sClampSizeButton2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sClampSizeButton3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sClampSizeButton4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sboxPlySeams_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sboxPlySeams_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call sUpdatePly.sUpdatePly(1)
End Sub
Private Sub sProjectNameBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sProjectTitleBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sProjectAddressBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sDateBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sSheetIssuedForBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sScaleBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sDrawnByBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sJobNoBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sSheetBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sSuffixBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sAreaBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sSaveButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub
Private Sub sLoadButton_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then ColumnCreator.Hide
End Sub

Private Sub sSaveButton_Click()
    Dim SheetDataFilePath As String
    Dim SheetDataFileName As String
    Dim SheetDataFilePathAndName As String
    Dim sHostName As String
    Dim fso As New FileSystemObject
    Dim oFile As Object
    
    'Set file name and path
    SheetDataFilePath = "C:\ProgramData\Autodesk\VBA Saved Files"
    SheetDataFileName = "Scissor Clamp Column Sheet Data.txt"
    SheetDataPathAndName = SheetDataFilePath & "\" & SheetDataFileName
    
    'Create folder path
    If Not FolderExists("C:\ProgramData") Then
        FolderCreate ("C:\ProgramData")
    End If
    If Not FolderExists("C:\ProgramData\Autodesk") Then
        FolderCreate ("C:\ProgramData\Autodesk")
    End If
    If Not FolderExists(SheetDataFilePath) Then
        FolderCreate (SheetDataFilePath)
    End If

    'Create a text file and write the following lines
    Set oFile = fso.CreateTextFile(SheetDataPathAndName, True) 'True allows overwriting
    
    oFile.WriteLine sProjectNameBox.Value
    oFile.WriteLine sProjectTitleBox.Value
    oFile.WriteLine sProjectAddressBox.Value
    oFile.WriteLine sDateBox.Value
    oFile.WriteLine sSheetIssuedForBox.Value
    oFile.WriteLine sScaleBox.Value
    oFile.WriteLine sDrawnByBox.Value
    oFile.WriteLine sJobNoBox.Value
    oFile.WriteLine sSheetBox.Value
    oFile.WriteLine sSuffixBox.Value
    oFile.WriteLine sAreaBox.Value
    
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing
End Sub
Public Sub sLoadButton_Click()
    Dim SheetDataFilePath As String
    Dim SheetDataFileName As String
    Dim SheetDataFilePathAndName As String
    Dim mystring As String
    Dim temp_lines(1 To 11) As String
    
    SheetDataFilePath = "C:\ProgramData\Autodesk\VBA Saved Files"
    SheetDataFileName = "Scissor Clamp Column Sheet Data.txt"
    SheetDataPathAndName = SheetDataFilePath & "\" & SheetDataFileName
    
    If Dir(SheetDataPathAndName) = "" Then
        MsgBox ("Error: You can't load a file you haven't saved")
        Exit Sub
    End If
    
    Open SheetDataPathAndName For Input As #1
    i = 1
    
    Do While Not EOF(1)
        Line Input #1, mystring
        temp_lines(i) = mystring
        i = i + 1
    Loop
    Close #1
    
    'Load the following lines to the corresponding variables
    sProjectNameBox.Value = temp_lines(1)
    sProjectTitleBox.Value = temp_lines(2)
    sProjectAddressBox.Value = temp_lines(3)
    sDateBox.Value = temp_lines(4)
    sSheetIssuedForBox.Value = temp_lines(5)
    sScaleBox.Value = temp_lines(6)
    sDrawnByBox.Value = temp_lines(7)
    sJobNoBox.Value = temp_lines(8)
    sSheetBox.Value = temp_lines(9)
    sSuffixBox.Value = temp_lines(10)
    sAreaBox.Value = temp_lines(11)
End Sub*/