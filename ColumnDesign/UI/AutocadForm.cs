/*
Private Sub Image3_Click()
    If ColumnCreator.lblAxis.Caption = "X" Then
        ColumnCreator.lblAxis.Caption = "Y"
    ElseIf ColumnCreator.lblAxis.Caption = "Y" Then
        ColumnCreator.lblAxis.Caption = "X"
    End If
    Call UpdatePly.UpdatePly(0) 'Update the plywood layout sketch
End Sub

Private Sub Label26_Click()
    If ColumnCreator.lblAxis.Caption = "X" Then
        ColumnCreator.lblAxis.Caption = "Y"
    ElseIf ColumnCreator.lblAxis.Caption = "Y" Then
        ColumnCreator.lblAxis.Caption = "X"
    End If
    Call UpdatePly.UpdatePly(0) 'Update the plywood layout sketch
End Sub

Private Sub lblAxis_Click()
    If ColumnCreator.lblAxis.Caption = "X" Then
        ColumnCreator.lblAxis.Caption = "Y"
    ElseIf ColumnCreator.lblAxis.Caption = "Y" Then
        ColumnCreator.lblAxis.Caption = "X"
    End If
    Call UpdatePly.UpdatePly(0) 'Update the plywood layout sketch
End Sub

Private Sub WinDim1_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    WinDim2.Value = ConvertFtIn(ConvertToNum(HeightBox.Value) - ConvertToNum(WinDim1.Value))
    Call UpdatePly.UpdatePly(2) 'Update the plywood layout sketch
End Sub
Private Sub WinDim2_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer)
    WinDim1.Value = ConvertFtIn(ConvertToNum(HeightBox.Value) - ConvertToNum(WinDim2.Value))
    Call UpdatePly.UpdatePly(2) 'Update the plywood layout sketch
End Sub


Private Sub CommandButton2_Click()
    'Call scissor clamp module
    Call CreateScissorClamp
End Sub

Private Sub CommandButton1_Click()

'#####################################################################################
'                                     T E X T
'#####################################################################################
         
    TODO Down
    
    Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt1, 100, _
    "• COLUMN SIZE = " & x & "'' X " & y & "''" & vbCrLf & vbCrLf & _
    "• NUMBER OF COLUMN FORMS = " & n_col & "-EA" & vbCrLf & vbCrLf & _
    "• COLUMN FORM WEIGHT (APPROXIMATE) = " & col_wt & "-LBS" & vbCrLf & vbCrLf & _
    "• PLYWOOD = 3/4'' PLYFORM (" & ply_name & "), CLASS-1 (MIN)" & vbCrLf & vbCrLf & _
    "• COLUMN FORMS AND CLAMP SPACING LAYOUTS FOR L4 X 3 X 1/4 GATES LOK-FAST COLUMN CLAMPS ARE DESIGNED FOR A POUR RATE = FULL LIQUID HEAD U.N.O." & vbCrLf & vbCrLf & _
    "• CONTACT THE MCC ENGINEER PRIOR TO ANY CHANGES OR MODIFICATIONS TO THE DETAILS ON THIS SHEET.")
    MTextObject1.StyleName = "Arial Narrow"
    MTextObject1.Height = 2
    
    'Plywood quantities
    For i = 1 To UBound(ply_cuts, 2)
        qty_text = qty_text & "• (" & ply_cuts(3, i) * n_col & "-EA) = (" & n_col & "-COL) X (" & ply_cuts(3, i) & "-EA/COL) @ " & ConvertFtIn(ply_cuts(1, i)) & " WIDE X " & ConvertFtIn(ply_cuts(2, i)) & " LONG 3/4'' PLYWOOD" & vbCrLf
    Next i
    qty_text = "PLYWOOD" & vbCrLf & qty_text & vbCrLf
    
    'Stud quantities
    If window = False Then
        qty_text = qty_text & "STUDS" & vbCrLf & "• (" & n_studs_total * n_col & "-EA) = (" & n_col & "-COL) X (" & n_studs_total & "-EA/COL) @ " & ConvertFtIn(z - 0.25) & " " & stud_name_full
    Else
        qty_text = qty_text & "STUDS" & vbCrLf & "• (" & (n_studs_total - n_studs_w) * n_col & "-EA) = (" & n_col & "-COL) X (" & (n_studs_total - n_studs_w) & "-EA/COL) @ " & ConvertFtIn(z - 0.25) & " " & stud_name_full & vbCrLf
        qty_text = qty_text & "• (" & n_studs_w * n_col & "-EA) = (" & n_col & "-COL) X (" & n_studs_w & "-EA/COL) @ " & ConvertFtIn(WinPos - WinStudOff - 0.25) & " " & stud_name_full & vbCrLf
        qty_text = qty_text & "• (" & n_studs_w * n_col & "-EA) = (" & n_col & "-COL) X (" & n_studs_w & "-EA/COL) @ " & ConvertFtIn(z - WinPos - WinStudOff) & " " & stud_name_full
    End If
    
    'Clamp quantities
    qty_text = qty_text & vbCrLf & vbCrLf & "COLUMN CLAMPS" & vbCrLf
    qty_text = qty_text & "• (" & n_clamps * n_col & "-EA) = (" & n_col & "-COL) X (" & n_clamps & "-EA/COL) @ GATES " & clamp_name & " LOK-FAST CLAMP ASSEMBLIES (SETS)."
    If n_reinf_angles > 0 Then
        qty_text = qty_text & vbCrLf & "• (" & n_reinf_angles * n_col & "-EA) = (" & n_col & "-COL) X (" & n_reinf_angles & "-EA/COL) @ GATES REINFORCING ANGLES"
    End If
    
    'Fastener quantities
    Dim n_nuts As Integer: n_nuts = 0 'Number of 1/2" coil rod nuts for threaded rods
    qty_text = qty_text & vbCrLf & vbCrLf & "FASTENERS" & vbCrLf & _
    "•   (" & n_bolts * n_col & "-EA) = (" & n_col & "-COL) X (" & n_bolts & "-EA/COL) @ 5/16'' X 3'' GATES FLAT HEAD BOLTS" & vbCrLf & _
    "•   (" & n_bolts * n_col & "-EA) = (" & n_col & "-COL) X (" & n_bolts & "-EA/COL) @ 5/16''-18 UNC NYLOK LOCK NUTS" & vbCrLf & _
    "•   (" & n_screws * n_col & "-EA) = (" & n_col & "-COL) X (" & n_screws & "-EA/COL) @ 1/4'' X 2-3/8'' GATES SPAX POWERLAG SCREWS" & vbCrLf
    
    '36" threaded rod if it occurs
    If n_top_clamps >= 2 Or SquaringCornerOptionPickingButton.Value = True Then
        qty_text = qty_text & _
        "•   (" & n_col * 2 & "-EA) = (" & n_col & "-COL) X (2-EA/COL) @ 1/2''-13 UNC X +/-36'' LONG ALL-THREADED ROD" & vbCrLf
        n_nuts = n_nuts + 8
    End If
    
    '14" threaded rod if it occurs (must have window, if there's a 36" rod only include a 14" rod if the 2nd top clamp is above the window, if there's no 36" rod always include a 14" rod)
    If (window = True And (n_top_clamps >= 2 Or SquaringCornerOptionPickingButton.Value = True) And clamp_spacing_con(2) > WinPos) Or (window = True And n_top_clamps < 2 And SquaringCornerOptionPickingButton.Value = False) Then
        qty_text = qty_text & _
        "•   (" & n_col & "-EA) = (" & n_col & "-COL) X (2-EA/COL) @ 1/2''-13 UNC X +/-14'' LONG ALL-THREADED ROD" & vbCrLf
        n_nuts = n_nuts + 4
    End If
    
    'Nuts and washers for threaded rods
    If n_nuts > 0 Then
        qty_text = qty_text & _
        "•   (" & n_nuts * n_col & "-EA) = (" & n_col & "-COL) X (" & n_nuts & "-EA/COL) @ 1/2''-13 UNC NYLOK LOCK NUTS" & vbCrLf & _
        "•   (" & n_nuts * n_col & "-EA) = (" & n_col & "-COL) X (" & n_nuts & "-EA/COL) @ 1/2'' STANDARD FLAT WASHER" & vbCrLf & vbCrLf
    End If
    
    'Chamfer quantity
    qty_text = qty_text & "3/4'' GATES PLASTIC CHAMFER (BASED ON 12' CHAMFER LENGTHS)" & vbCrLf & _
    "•   (" & n_col * n_chamf & "-EA) = (" & n_col & "-COL) X (" & n_chamf & "-EA/COL) @ " & chamf_length & "'-0'' LONG PIECES"
    
    'AFB quantity
    qty_text = qty_text & vbCrLf & vbCrLf & "GATES ADJUSTABLE FORM BRACES (INCLUDING FORM BASE PLATES)" & vbCrLf & _
    "•   (" & n_col * 3 & "-EA) = (" & n_col & "-COL) X (3EA/COL) " & brace_name & " LONG GATES AFB" & vbCrLf & vbCrLf
    
    'Hoisting slings
    If col_wt <= 2400 Then
        qty_text = qty_text & "HOISTING SLINGS" & vbCrLf & _
        "•   (" & n_col * 2 & "-EA) = (" & n_col & "-COL) X (2EA/COL) ENDLESS ROUND SLINGS; LIFTEX P/N ''ENR1'', PURPLE." & vbCrLf & "SWL = 2400-LBS. PER SLING IN CHOKER CONFIGURATION."
    Else
        qty_text = qty_text & "HOISTING SLINGS" & vbCrLf & _
        "•   (" & n_col * 2 & "-EA) = (" & n_col & "-COL) X (2EA/COL) ENDLESS ROUND SLINGS; LIFTEX P/N ''ENR2'', GREEN." & vbCrLf & "SWL = 4800-LBS. PER SLING IN CHOKER CONFIGURATION."
    End If
        
    'Set all count and fab note texts
    Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt3, 100, qty_text)
    MTextObject1.Height = 2
    MTextObject1.StyleName = "Arial Narrow" 'Forces use of Arial Narrow text style
        
    '#####################################################################################
    '                                 D R A W I N G
    '#####################################################################################

    Call DrawRectangle(ptE(0) - ply_thk, ptE(1) + stud_base_gap, ptE(0) - ply_thk - 1.5, ptE(1) + z)
    
    'Draw studs
        'Side A
        For i = 1 To n_studs_x
            pt6(0) = ptA(0) + ply_width_x + chamf_thk - 3.5 - stud_spacing_x(i)
            pt6(1) = ptA(1) + stud_base_gap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
            Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)
        Next i
        
        'Side B
        If DrawB = True Then
            For i = 1 To n_studs_y
                pt6(0) = ptB(0) + ply_width_y + chamf_thk - 3.5 - stud_spacing_y(i)
                pt6(1) = ptB(1) + stud_base_gap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
                Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)
            Next i
        End If
        
        'Side W
        If DrawW = True Then
            For i = 1 To n_studs_w
                pt6(0) = ptW(0) + ply_width_w + chamf_thk - 3.5 - stud_spacing_w(i): pt6(1) = ptW(1) + stud_base_gap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs below window
                Call ChangeProp(BlockRefObj, "Distance1", WinPos - WinStudOff - stud_base_gap)
                pt6(1) = ptW(1) + WinPos + WinStudOff + WinGap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs on window
                Call ChangeProp(BlockRefObj, "Distance1", z - WinPos - WinStudOff)
            Next i
        End If
        
    'Draw plywood sheets
    ThisDrawing.ActiveLayer = ThisDrawing.Layers("PLYA")
    pt4(1) = ptA(1)
    For i = 1 To UBound(ply_seams)
        pt4(0) = ptA(0)
        pt4(1) = pt4(1) + ply_seams(i)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt4, "VBA_PLY_SHEET", 1#, 1#, 1#, 0) 'Insert plywood sheets
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
        Call ChangeProp(BlockRefObj, "Distance2", ply_seams(i))
        
        If DrawB = True Then
            pt4(0) = ptB(0)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt4, "VBA_PLY_SHEET", 1#, 1#, 1#, 0) 'Insert plywood sheets
            Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
            Call ChangeProp(BlockRefObj, "Distance2", ply_seams(i))
        End If
    Next i
    If DrawW = True Then
        pt4(0) = ptW(0): pt4(1) = ptW(1)
        For i = 1 To UBound(ply_seams_win)
            pt4(1) = pt4(1) + ply_seams_win(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt4, "VBA_PLY_SHEET", 1#, 1#, 1#, 0) 'Insert plywood sheets
            Call ChangeProp(BlockRefObj, "Distance1", ply_width_w)
            Call ChangeProp(BlockRefObj, "Distance2", ply_seams_win(i))
            If Round(pt4(1) - ptW(1), 3) = Round(WinPos, 3) Then pt4(1) = pt4(1) + WinGap 'Create a visual gap in the plywood where the window seam occurs
        Next i
    End If
    
    ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
        'Draw screw heads
            'Side A
            ply_tot = 0
            For i = 1 To UBound(ply_seams)
                ply_tot = ply_tot + ply_seams(i)
                For j = 1 To UBound(stud_spacing_x)
                    pt29(0) = ptA(0) + ply_width_x + chamf_thk - stud_spacing_x(j) - 1.75: pt29(1) = ptA(1) + ply_tot - 2 'set a screw head below the ply seam for each location on ply_seams
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                    If i <> UBound(ply_seams) Then
                        pt29(0) = ptA(0) + ply_width_x + chamf_thk - stud_spacing_x(j) - 1.75: pt29(1) = ptA(1) + ply_tot + 2 'set a screw head above the ply seam for each loc except the top
                        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                    End If
                Next j
            Next i
            For j = 1 To UBound(stud_spacing_x) 'set screw heads at the bottom
                pt29(0) = ptA(0) + ply_width_x + chamf_thk - stud_spacing_x(j) - 1.75: pt29(1) = ptA(1) + 2
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
            Next j
            'Add the text notes on the right if no side B
            If DrawB = False Then
                pt_blk(0) = ptA(0) + ply_width_x + chamf_thk - stud_spacing_x(1) - 1.75: pt_blk(1) = ptA(1) + ply_seams(1) - 2
                If ply_seams(1) < 48 And UBound(ply_seams) > 1 Then pt_blk(1) = pt_blk(1) + ply_seams(2) 'Move note up if ply_seams(1) is small enough that it overlaps with the other text note
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES_SCREWS", 1#, 1#, 1#, 0)
                
                pt_blk(0) = ptA(0) + ply_width_x: pt_blk(1) = ptA(1)
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES", 1#, 1#, 1#, 0)
            End If
            
            'Side B
            If DrawB = True Then
                ply_tot = 0
                For i = 1 To UBound(ply_seams)
                    ply_tot = ply_tot + ply_seams(i)
                    For j = 1 To UBound(stud_spacing_y)
                        pt29(0) = ptB(0) + ply_width_y + chamf_thk - stud_spacing_y(j) - 1.75: pt29(1) = ptB(1) + ply_tot - 2 'set a screw head below the ply seam for each location on ply_seams
                        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                        If i <> UBound(ply_seams) Then
                            pt29(0) = ptB(0) + ply_width_y + chamf_thk - stud_spacing_y(j) - 1.75: pt29(1) = ptB(1) + ply_tot + 2 'set a screw head above the ply seam for each loc except the top
                            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                        End If
                    Next j
                Next i
                For j = 1 To UBound(stud_spacing_y) 'set screw heads at the bottom
                    pt29(0) = ptB(0) + ply_width_y + chamf_thk - stud_spacing_y(j) - 1.75: pt29(1) = ptB(1) + 2
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                Next j
                'Add the text notes on the right if no window, left side if there is a window and a B side
                If DrawW = False Then 'Right side
                    pt_blk(0) = ptB(0) + ply_width_y + chamf_thk - stud_spacing_y(1) - 1.75: pt_blk(1) = ptB(1) + ply_seams(1) - 2
                    If ply_seams(1) < 48 And UBound(ply_seams) > 1 Then pt_blk(1) = pt_blk(1) + ply_seams(2) 'Move note up if ply_seams(1) is small enough that it overlaps with the other text note
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES_SCREWS", 1#, 1#, 1#, 0)
                    pt_blk(0) = ptB(0) + ply_width_y: pt_blk(1) = ptB(1)
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES", 1#, 1#, 1#, 0)
                ElseIf DrawW = True Then 'Left side
                    pt_blk(0) = ptB(0) + ply_width_y + chamf_thk - stud_spacing_y(UBound(stud_spacing_y)) - 1.75: pt_blk(1) = ptB(1) + ply_seams(1) - 2
                    If ply_seams(1) < 48 And UBound(ply_seams) > 1 Then pt_blk(1) = pt_blk(1) + ply_seams(2) 'Move note up if ply_seams(1) is small enough that it overlaps with the other text note
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES_SCREWS_MIRRORED", 1#, 1#, 1#, 0)
                    pt_blk(0) = ptB(0) + ply_width_y: pt_blk(1) = ptB(1)
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_ELEVATION_NOTES_MIRRORED", 1#, 1#, 1#, 0)
                    Call ChangeProp(BlockRefObj, "Distance1", y + 12)
                End If
            End If
            
            'Side W
            If DrawW = True Then
                ply_tot = 0
                For i = 1 To UBound(ply_seams_win)
                    ply_tot = ply_tot + ply_seams_win(i)
                    For j = 1 To UBound(stud_spacing_w)
                        pt29(0) = ptW(0) + ply_width_w + chamf_thk - stud_spacing_w(j) - 1.75: pt29(1) = ptW(1) + ply_tot - 2 'set a screw head below the ply seam for each location on ply_seams_win
                        If Round(ply_tot, 3) = Round(WinPos, 3) Then pt29(1) = pt29(1) - WinStudOff 'Move screws down where studs are held back at the window seam
                        If Round(ply_tot, 3) > Round(WinPos, 3) Then pt29(1) = pt29(1) + WinGap
                        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                        If i <> UBound(ply_seams_win) Then
                            pt29(0) = ptW(0) + ply_width_w + chamf_thk - stud_spacing_w(j) - 1.75: pt29(1) = ptW(1) + ply_tot + 2 'set a screw head above the ply seam for each loc except the top
                            If Round(ply_tot, 3) = Round(WinPos, 3) Then
                                pt29(1) = pt29(1) + WinStudOff + WinGap 'Move screws up where studs are held back at the window seam
                            End If
                            If Round(ply_tot, 3) > Round(WinPos, 3) Then pt29(1) = pt29(1) + WinGap
                            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                        End If
                    Next j
                Next i
                For j = 1 To UBound(stud_spacing_w) 'set screw heads at the bottom
                    pt29(0) = ptW(0) + ply_width_w + chamf_thk - stud_spacing_w(j) - 1.75: pt29(1) = ptW(1) + 2
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt29, "VBA_SCREW_HEAD", 1#, 1#, 1#, 0) 'Insert screw heads
                Next j
            End If

        
        'Draw chamfer
        pt18(0) = ptA(0) + ply_width_x: pt18(1) = ptA(1) + stud_base_gap
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt18, "VBA_CHAMF", 1#, 1#, 1#, 0) 'Insert chamfer
        Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)
        
        If DrawB = True Then
            pt18(0) = ptB(0) + ply_width_y: pt18(1) = ptB(1) + stud_base_gap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt18, "VBA_CHAMF", 1#, 1#, 1#, 0) 'Insert chamfer
            Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)

        End If

        If DrawW = True Then
            pt18(0) = ptW(0) + ply_width_w: pt18(1) = ptW(1) + stud_base_gap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt18, "VBA_CHAMF", 1#, 1#, 1#, 0) 'Insert chamfer below window
            Call ChangeProp(BlockRefObj, "Distance1", WinPos - stud_base_gap)
            pt18(0) = ptW(0) + ply_width_w: pt18(1) = ptW(1) + WinPos + WinGap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt18, "VBA_CHAMF", 1#, 1#, 1#, 0) 'Insert chamfer on window
            Call ChangeProp(BlockRefObj, "Distance1", z - WinPos)
        End If
        
    'Draw elevation view (clamps) specifc items
        'Studs
        If window = False Then
            For i = 1 To n_studs_e
                pt20(0) = ptE(0) + stud_spacing_e(i): pt20(1) = ptE(1) + stud_base_gap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
                Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap): Call ChangeProp(BlockRefObj, "Visibility1", "Solid")
            Next i
        ElseIf window = True Then
            If DrawW = True Then
                For i = 1 To n_studs_e
                    pt20(0) = ptE(0) + stud_spacing_e(i): pt20(1) = ptE(1) + stud_base_gap
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
                    Call ChangeProp(BlockRefObj, "Distance1", WinPos - stud_base_gap - WinStudOff): Call ChangeProp(BlockRefObj, "Visibility1", "Solid")
                    pt21(0) = ptE(0) + stud_spacing_e(i): pt21(1) = ptE(1) + WinPos + WinStudOff
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt21, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
                    Call ChangeProp(BlockRefObj, "Distance1", z - WinPos - WinStudOff): Call ChangeProp(BlockRefObj, "Visibility1", "Solid")
                Next i
            End If
        End If
        
        'Plywood
        'Outer ply boundry of elevation view
        ThisDrawing.ActiveLayer = ThisDrawing.Layers("PLYA")
        Call DrawRectangle(ptE(0) + chamf_thk, ptE(1), ptE(0) + chamf_thk + ply_width_e, ptE(1) + z)
        
        'Plywood edge view
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(ptE, "VBA_PLY", 1#, 1#, 1#, pi / 2)
        Call ChangeProp(BlockRefObj, "Distance1", z)
        
        'Plywood seam at window
        If window = True Then
            pt20(0) = ptE(0) + chamf_thk: pt20(1) = ptE(1) + WinPos
            pt21(0) = ptE(0) + ply_width_e + chamf_thk: pt21(1) = ptE(1) + WinPos
            Set LineObj = ThisDrawing.ModelSpace.AddLine(pt20, pt21)
        End If
        
        'Draw 2x2s for window lock
        If window = True Then
            ThisDrawing.ActiveLayer = ThisDrawing.Layers("0")
            If z - WinPos < 18 Then
                pt20(1) = ptE(1) + z
            Else
                pt20(1) = ptE(1) + WinPos + 18
            End If
            pt21(1) = pt20(1) - 36
            
            Dim Inserted2x2Note As Boolean: Inserted2x2Note = False
            For i = 1 To UBound(stud_spacing_w) - 1
                If stud_spacing_w(i + 1) - stud_spacing_w(i) - 3.5 > min_2x2_gap Then
                    pt20(0) = ptE(0) + stud_spacing_w(i) + 3.5
                    pt21(0) = pt20(0) + 1.5:
                    Call DrawRectangle(pt20(0), pt20(1), pt21(0), pt21(1))
                    
                    'Add text note under first clamp below window
                    If Inserted2x2Note = False Then
                        Inserted2x2Note = True
                        For j = 1 To n_clamps
                            If clamp_spacing_con(j) < WinPos Then
                                pt18(0) = ptE(0) + stud_spacing_w(UBound(stud_spacing_w) - 1) + 3.5 + 1.5
                                pt18(1) = ptE(1) + clamp_spacing_con(j) - 6
                                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt18, "VBA_2X2_NOTES", 1#, 1#, 1#, 0)
                                GoTo NoteInserted
                            End If
                        Next j
                    End If
NoteInserted:
                End If
            Next i
        End If
        
        'Clamps and dims
        ThisDrawing.ActiveLayer = ThisDrawing.Layers("FMTEXTA")
        pt21(0) = ptE(0) + chamf_thk + ply_width_e: pt21(1) = ptE(1) + z
        Call DrawDimLin(pt21(0) - clamp_L + 3.75, pt21(1) - clamp_spacing(1), ptE(0) - ply_thk, pt21(1), pt21(0) - clamp_L + 3.75 - 4, (pt21(1) * 2 - clamp_spacing(1)) / 2, pi / 2)     'Add top dim
        
        For i = 1 To n_clamps
            'Draw clamps and dims
            pt21(1) = pt21(1) - clamp_spacing(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt21, clamp_block_pr, 1#, 1#, 1#, 0) 'Insert clamp profile views
            Call ChangeProp(BlockRefObj, "Distance1", 1.5 + ply_thk + chamf_thk + ply_width_e)
            If i <> 1 Then
                Call DrawDim(pt21(0) - clamp_L + 3.75, pt21(1), pt21(0) - clamp_L + 3.75, pt21(1) + clamp_spacing(i), pt21(0) - clamp_L + 3.75 - 4, (pt21(1) * 2 + clamp_spacing(i)) / 2)
            End If
            
            'Draw inches from bottom
            pt22(0) = pt21(0) - clamp_L + 3.75 - 12: pt22(1) = ptE(1) + clamp_spacing_con(i) + 1.5
            clamp_str = CStr(clamp_spacing_con(i)) & """"
            Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt22, 9, clamp_str)
            MTextObject2.Height = 2.5
            
            'Draw "CLAMP" and "TOP CLAMP" labels
            pt25(0) = ptE(0) + chamf_thk + ply_width_e + 9: pt25(1) = pt21(1) + 5.75 - 3 * Abs(WinY = True)
            
            If clamp_size = 1 And ply_width_e - 1.5 >= 19 Then
                pt25(1) = pt25(1) - 3
            ElseIf clamp_size = 2 And ply_width_e - 1.5 >= 31 Then
                pt25(1) = pt25(1) - 3
            ElseIf clamp_size = 3 And ply_width_e - 1.5 >= 43 Then
                pt25(1) = pt25(1) - 3
            End If
            
            If i <= n_top_clamps Then
                Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt25, 24, "TOP CLAMP")
            Else
                Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt25, 17, "CLAMP")
            End If
            MTextObject2.AttachmentPoint = acAttachmentPointTopLeft
            MTextObject2.Height = 2.5
        Next i
        Call DrawDimLin(pt21(0) - clamp_L + 3.75, ptE(1) + bot_clamp_gap, pt21(0) - x - 8, ptE(1), pt21(0) - clamp_L + 3.75 - 4, (2 * ptE(1) + bot_clamp_gap) / 2, pi / 2) 'Add bottom dim
        
        'Draw 36" coil rod and nuts connecting top 2 clamps if column is tall enough
        If n_top_clamps >= 2 Or SquaringCornerOptionPickingButton.Value = True Then
            pt27(0) = ptE(0) + chamf_thk + 2: pt27(1) = ptE(1) + z - clamp_spacing(1) + 2.25 - 0.5 * Abs(WinY = True)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD", 1#, 1#, 1#, pi)
            If ply_width_w >= 37.5 Then
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD_NOTES", 1#, 1#, 1#, 0)
            ElseIf SquaringCornerOptionPickingButton.Value = True Then
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD_NOTES_C", 1#, 1#, 1#, 0)
                Call ChangeProp(BlockRefObj, "Distance1", Abs(40 - (36 - ply_width_e)))
            Else
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD_NOTES_B", 1#, 1#, 1#, 0)
            End If
            pt27(1) = pt27(1) - 2
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, 0)
            pt27(1) = pt27(1) - 0.25
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, pi)
            pt27(1) = pt27(1) - clamp_spacing(2)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, pi)
            pt27(1) = pt27(1) + 0.25
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, 0)
        End If
        
        'Draw 14" coil rod and nuts connecting clamps around window seam
        If window = True Then
            'If there's a 36" rod, check if it's bolting the same clamps as the 14" rod.
            If n_top_clamps >= 2 Or SquaringCornerOptionPickingButton.Value = True Then
                'Don't add this if the 36" coil rod is also bolting the window clamps
                If clamp_spacing_con(2) > WinPos Then
Draw14InRodAnways:
                    pt27(0) = ptE(0) + chamf_thk + ply_width_w / 2: pt27(1) = ptE(1) + WinPos + win_clamp_top_max + 2.625 - 0.5 * Abs(WinY = True)
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD_14", 1#, 1#, 1#, pi)
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_COIL_ROD_NOTES_14", 1#, 1#, 1#, 0)
                    Call ChangeProp(BlockRefObj, "Distance1", Abs(clamp_spacing_con(1) - clamp_spacing_con(2) - 12))
                    pt27(1) = pt27(1) - 2.375
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, 0)
                    pt27(1) = pt27(1) - 0.25
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, pi)
                    pt27(1) = pt27(1) - win_clamp_top_max - win_clamp_bot_max + 0.25
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, 0)
                    pt27(1) = pt27(1) - 0.25
                    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt27, "VBA_NUT_5-8", 1#, 1#, 1#, pi)
                End If
            Else
                GoTo Draw14InRodAnways
            End If
        End If
        
        'Draw reinforcing angle on all clamps at or below 87" from top of column on elevation veiw.
        If x > 40 Then
            For i = 1 To UBound(clamp_spacing_con) - 1 'Last value is a 0
                If (z - clamp_spacing_con(i)) >= 87 Then
                    If WinY = True Then 'Flip and raise reinforcing angle when displayed on Y side of clamp in elevation view
                        pt31(0) = ptE(0) + chamf_thk + (ply_width_e - 36) / 2 + 36
                        pt31(1) = ptE(1) + clamp_spacing_con(i) - 0.25
                        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt31, "VBA_REINFORCING_ANGLE", 1#, 1#, 1#, pi)
                    Else
                        pt31(0) = ptE(0) + chamf_thk + (ply_width_e - 36) / 2
                        pt31(1) = ptE(1) + clamp_spacing_con(i)
                        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt31, "VBA_REINFORCING_ANGLE", 1#, 1#, 1#, 0)
                    End If
                End If
            Next i
        End If
        
        'Lifting slings
        If SquaringCornerOptionRegularButton.Value = True Then
            pt26(0) = ptE(0) - ply_thk - 1.25: pt26(1) = ptE(1) + z - clamp_spacing(1)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt26, "VBA_LIFTING_SLING_A", 1#, 1#, 1#, 0) 'Insert left sling
            pt26(0) = ptE(0) + chamf_thk + ply_width_e + 1.3: pt26(1) = ptE(1) + z - clamp_spacing(1) + 3
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt26, "VBA_LIFTING_SLING_B", 1#, 1#, 1#, 0) 'Insert right sling
        ElseIf SquaringCornerOptionPickingButton.Value = True Then
            pt26(0) = ptE(0) - ply_thk - 1.5: pt26(1) = ptE(1) + z - clamp_spacing(1) - 0.25
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt26, "VBA_LIFTING_CORNER_A", 1#, 1#, 1#, 0) 'Insert left sling
            pt26(0) = ptE(0) + chamf_thk + ply_width_e: pt26(1) = ptE(1) + z - clamp_spacing(1) - 0.25
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt26, "VBA_LIFTING_CORNER_B", 1#, 1#, 1#, 0) 'Insert right sling
        End If
        
        'AFB
        pt28(0) = ptE(0) + chamf_thk + ply_width_e + 4: pt28(1) = ptE(1) + clamp_spacing_con(brace_clamp)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, "VBA_AFB_SIDE", 1#, 1#, 1#, 0) 'Insert AFB side view
        
        pt28(0) = pt28(0) + 1.9375: pt28(1) = pt28(1) + 0.6789
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, "VBA_BRACE_ANGLE", 1#, 1#, 1#, 0) 'Insert brace angle
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, brace_block, 1#, 1#, 1#, 0) 'Insert brace dynamic block
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, "VBA_BRACE_NOTES", 1#, 1#, 1#, 0) 'Insert brace notes
        
        'Insert safety chain block and note, stretch note if the block is placed too low and will conflict with "VBA_ELEVATION_NOTES_MIRRORED"
        pt28(0) = pt28(0) - 5.9375: pt28(1) = ptE(1) + clamp_spacing_con(chain_clamp)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, "VBA_CHAIN", 1#, 1#, 1#, 0)
        If pt28(1) - pt_o(1) < 62 Then
            Call ChangeProp(BlockRefObj, "Distance1", 6.27 + 62 - (pt28(1) - pt_o(1)))
        End If
        
        pt28(0) = ptE(0) + 4.68: pt28(1) = ptE(1) + clamp_spacing_con(brace_clamp)
        If window = False Then
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt28, "VBA_AFB_FACE", 1#, 1#, 1#, 0) 'Insert AFB face view
        End If
        
        ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")

        'Base framing lumber
        pt20(0) = ptE(0) - ply_thk - 1.5 - 3.5: pt20(1) = ptE(1)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4", 1#, 1#, 1#, 0) 'Insert studs
        pt20(0) = ptE(0) + chamf_thk + ply_width_e: pt20(1) = ptE(1)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4", 1#, 1#, 1#, 0) 'Insert studs
        pt20(0) = ptE(0) - ply_thk - 1.5 - 6: pt20(1) = ptE(1) + 3
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4_SIDE", 1#, 1#, 1#, 0) 'Insert studs
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_e + chamf_thk + ply_thk + 1.5 + 12)
        pt30(0) = ptE(0) + chamf_thk + ply_width_e + 6: pt30(1) = ptE(1) + 2.25
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_DOWN_PLATE_NOTES", 1#, 1#, 1#, 0)
    
    'Draw panel sections
        'Side A
        pt8(0) = ptA(0) + ply_width_x: pt8(1) = ptA(1) + z + 18 + Abs(window) * WinGap
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt8, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
        For i = 1 To n_studs_x
            pt11(0) = ptA(0) + ply_width_x + chamf_thk - 3.5 - stud_spacing_x(i)
            pt11(1) = ptA(1) + z + 18 + Abs(window) * WinGap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
        Next i
        
        'Side B
        If DrawB = True Then
            pt9(0) = ptB(0) + ply_width_y: pt9(1) = ptB(1) + z + 18 + Abs(window) * WinGap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi) 'Insert ply with chamfer
            Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
            For i = 1 To n_studs_y
                pt11(0) = ptB(0) + ply_width_y + chamf_thk - 3.5 - stud_spacing_y(i)
                pt11(1) = ptB(1) + z + 18 + Abs(window) * WinGap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
            Next i
        End If

        'Side W
        If DrawW = True Then
            pt10(0) = ptW(0) + ply_width_w: pt10(1) = ptW(1) + z + 18 + Abs(window) * WinGap
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt10, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi) 'Insert ply with chamfer
            Call ChangeProp(BlockRefObj, "Distance1", ply_width_w)
            For i = 1 To n_studs_w
                pt11(0) = ptW(0) + ply_width_w + chamf_thk - 3.5 - stud_spacing_w(i)
                pt11(1) = ptW(1) + z + 18 + Abs(window) * WinGap
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
            Next i
        End If
        
        'Dim sections
            'Side A
            Call DrawDim(pt8(0), pt8(1) - ply_thk, pt8(0) - ply_width_x, pt8(1) - ply_thk, (pt8(0) - ply_width_x - pt8(0)) / 2, pt8(1) - ply_thk - 3) 'Dimension plywood
            Call DrawDimLin(pt8(0) + chamf_thk, pt8(1) - ply_thk - 0.5, pt8(0) - ply_width_x, pt8(1) - ply_thk, (pt8(0) - ply_width_x - pt8(0)) / 2, pt8(1) - ply_thk - 6.75, 0) 'Dimension overall panel
            For i = 1 To n_studs_x
                Call DrawDimLin(pt8(0) + chamf_thk - stud_spacing_x(UBound(stud_spacing_x) - i + 1), pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) + chamf_thk - stud_spacing_x(UBound(stud_spacing_x) - i + 1))) / 2, pt8(1) + 4 + i * 4, 0) 'Dimension studs from rightmost (in plan) stud
            Next i
            Call DrawDimLin(pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5 + chamf_thk, pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5 + chamf_thk)) / 2, pt8(1) + 4, 0) 'Dimension leftmost (in plan) stud to face of ply
            Call DrawDimLin(pt8(0) + chamf_thk - stud_start_offset, pt8(1) + 1.5, pt8(0) + chamf_thk, pt8(1), ((pt8(0) + chamf_thk) - (pt8(0) + chamf_thk - stud_start_offset)) / 2, pt8(1) + 4 * (n_studs_x + 1), 0) 'dimension stud_start_gap
            
            'Side B
            If DrawB = True Then
                Call DrawDim(pt9(0), pt9(1) - ply_thk, pt9(0) - ply_width_y, pt9(1) - ply_thk, (pt9(0) - ply_width_y - pt9(0)) / 2, pt9(1) - ply_thk - 3) 'Dimension plywood
                Call DrawDimLin(pt9(0) + chamf_thk, pt9(1) - ply_thk - 0.5, pt9(0) - ply_width_y, pt9(1) - ply_thk, (pt9(0) - ply_width_y - pt9(0)) / 2, pt9(1) - ply_thk - 6.75, 0) 'Dimension overall panel
                For i = 1 To n_studs_y
                    Call DrawDimLin(pt9(0) + chamf_thk - stud_spacing_y(UBound(stud_spacing_y) - i + 1), pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) + chamf_thk - stud_spacing_y(UBound(stud_spacing_y) - i + 1))) / 2, pt9(1) + 4 + i * 4, 0) 'Dimension studs from rightmost (in plan) stud
                Next i
                Call DrawDimLin(pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5 + chamf_thk, pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5 + chamf_thk)) / 2, pt9(1) + 4, 0) 'Dimension leftmost (in plan) stud to face of ply
                Call DrawDimLin(pt9(0) + chamf_thk - stud_start_offset, pt9(1) + 1.5, pt9(0) + chamf_thk, pt9(1), ((pt9(0) + chamf_thk) - (pt9(0) + chamf_thk - stud_start_offset)) / 2, pt9(1) + 4 * (n_studs_y + 1), 0) 'dimension stud_start_gap

                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_TOP_SECTION_DETAILS1", 1#, 1#, 1#, 0) 'Insert text notes, blow up detail at top section
            End If
            
            'Side W
            If DrawW = True Then
                Call DrawDim(pt10(0), pt10(1) - ply_thk, pt10(0) - ply_width_w, pt10(1) - ply_thk, (pt10(0) - ply_width_w - pt10(0)) / 2, pt10(1) - ply_thk - 3) 'Dimension plywood
                Call DrawDimLin(pt10(0) + chamf_thk, pt10(1) - ply_thk - 0.5, pt10(0) - ply_width_w, pt10(1) - ply_thk, (pt10(0) - ply_width_w - pt10(0)) / 2, pt10(1) - ply_thk - 6.75, 0) 'Dimension overall panel
                For i = 1 To n_studs_w
                    Call DrawDimLin(pt10(0) + chamf_thk - stud_spacing_w(UBound(stud_spacing_w) - i + 1), pt10(1) + 1.5, pt10(0) - ply_width_w, pt10(1), ((pt10(0) - ply_width_w) - (pt10(0) + chamf_thk - stud_spacing_w(UBound(stud_spacing_w) - i + 1))) / 2, pt10(1) + 4 + i * 4, 0) 'Dimension studs from rightmost (in plan) stud
                Next i
                Call DrawDimLin(pt10(0) - stud_spacing_w(UBound(stud_spacing_w)) - 3.5 + chamf_thk, pt10(1) + 1.5, pt10(0) - ply_width_w, pt10(1), ((pt10(0) - ply_width_w) - (pt10(0) - stud_spacing_w(UBound(stud_spacing_w)) - 3.5 + chamf_thk)) / 2, pt10(1) + 4, 0) 'Dimension leftmost (in plan) stud to face of ply
                Call DrawDimLin(pt10(0) + chamf_thk - stud_start_offset, pt10(1) + 1.5, pt10(0) + chamf_thk, pt10(1), ((pt10(0) + chamf_thk) - (pt10(0) + chamf_thk - stud_start_offset)) / 2, pt10(1) + 4 * (n_studs_w + 1), 0) 'dimension stud_start_gap
            End If
            
        'Dim plywood
            'Side A
            PlyTemp = ptA(1) + ply_seams(1)
            Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - ptA(1)) / 2, " PLYWOOD")
            If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
                For i = 2 To UBound(ply_seams)
                    PlyTemp = PlyTemp + ply_seams(i)
                    Call DrawDimSuffix(ptA(0) + ply_width_x, PlyTemp - ply_seams(i), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
                Next i
            End If
            Call DrawDimSuffix(ptA(0), ptA(1), ptA(0), ptA(1) + z, ptA(0) - 9, (2 * ptA(1) + z) / 2, " OVERALL HEIGHT")
            Call DrawDimSuffix(ptA(0), ptA(1) + z, ptA(0), ptA(1) + stud_base_gap, ptA(0) - 4.5, ((ptA(1) + z) + (ptA(1) + stud_base_gap)) / 2, " STUD")
            Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1) + z, ptA(0), ptA(1) + z, (2 * ptA(0) + ply_width_x) / 2, ptA(1) + z + 4, " PLYWOOD")
            
            'Side B
            If DrawB = True Then
                PlyTemp = ptB(1) + ply_seams(1)
                Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - ptB(1)) / 2, " PLYWOOD")
                If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
                    For i = 2 To UBound(ply_seams)
                        PlyTemp = PlyTemp + ply_seams(i)
                        Call DrawDimSuffix(ptB(0) + ply_width_y, PlyTemp - ply_seams(i), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
                    Next i
                End If
                Call DrawDimSuffix(ptB(0), ptB(1), ptB(0), ptB(1) + z, ptB(0) - 9, (2 * ptB(1) + z) / 2, " OVERALL HEIGHT")
                Call DrawDimSuffix(ptB(0), ptB(1) + z, ptB(0), ptB(1) + stud_base_gap, ptB(0) - 4.5, ((ptB(1) + z) + (ptB(1) + stud_base_gap)) / 2, " STUD")
                Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1) + z, ptB(0), ptB(1) + z, (2 * ptB(0) + ply_width_y) / 2, ptB(1) + z + 4, " PLYWOOD")
            End If
            
            'Side W
            If DrawW = True Then
                PlyTemp = ptW(1) + ply_seams_win(1)
                Call DrawDimSuffix(ptW(0) + ply_width_w, ptW(1), ptW(0) + ply_width_w, PlyTemp, ptW(0) + ply_width_w + 6, (PlyTemp - ptW(1)) / 2, " PLYWOOD") 'Dim lowest (first) plywood sheet
                If UBound(ply_seams_win) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
                    For i = 2 To UBound(ply_seams_win)
                        
                        If Round(PlyTemp - ptW(1), 3) = Round(WinPos, 3) Then PlyTemp = PlyTemp + WinGap 'Add the window gap distance when it's reached
                        PlyTemp = PlyTemp + ply_seams_win(i)
                        
                        'Pop out dimension if less than 26"
                        If Round(ply_seams_win(i), 3) < 26 Then
                            Call DrawDimLinSuffixLeader(ptW(0) + ply_width_w, PlyTemp - ply_seams_win(i), ptW(0) + ply_width_w, PlyTemp, ptW(0) + ply_width_w + 6, (PlyTemp + (PlyTemp - ply_seams_win(i))) / 2, ptW(0) + ply_width_w + 11, 4 + (PlyTemp + (PlyTemp - ply_seams_win(i))) / 2, " PLYWOOD", pi / 2)
                        Else
                            Call DrawDimSuffix(ptW(0) + ply_width_w, PlyTemp - ply_seams_win(i), ptW(0) + ply_width_w, PlyTemp, ptW(0) + ply_width_w + 6, (PlyTemp + (PlyTemp - ply_seams_win(i))) / 2, " PLYWOOD")
                        End If
                        
                    Next i
                End If
                Call DrawDimSuffix(ptW(0), ptW(1) + WinPos - WinStudOff, ptW(0), ptW(1) + stud_base_gap, ptW(0) - 4.5, (ptW(1) + WinPos - WinStudOff + ptW(1) + stud_base_gap) / 2, " STUD") 'lower stud
                Call DrawDimSuffix(ptW(0), ptW(1) + WinPos + WinStudOff + WinGap, ptW(0), ptW(1) + z + WinGap, ptW(0) - 4.5, (ptW(1) + z + ptW(1) + WinPos + WinStudOff + 2 * WinGap) / 2, " STUD") 'window stud
                Call DrawDimLinLeader(ptW(0), ptW(1) + WinPos - WinStudOff, ptW(0), ptW(1) + WinPos, ptW(0) - 4.5, (ptW(1) + WinPos - WinStudOff + ptW(1) + WinPos) / 2, ptW(0) - 8.5, ptW(1) + WinPos - 5.5, pi / 2) 'Lower stud offset dimension
                Call DrawDimLinLeader(ptW(0), ptW(1) + WinPos + WinStudOff + WinGap, ptW(0), ptW(1) + WinPos + WinGap, ptW(0) - 4.5, (ptW(1) + WinPos + WinStudOff + ptW(1) + WinPos + 2 * WinGap) / 2, ptW(0) - 8.5, ptW(1) + WinPos + WinGap + 5.5, pi / 2) 'window stud offset dimension
                Call DrawDimSuffix(ptW(0) + ply_width_w, ptW(1) + z + WinGap, ptW(0), ptW(1) + z + WinGap, (2 * ptW(0) + ply_width_w) / 2, ptW(1) + z + 4 + WinGap, " PLYWOOD")
            End If

            'Dim window seam to clamps on elevation view
            If window = True Then
                Call DrawDimLin(ptE(0) + chamf_thk, ptE(1) + WinPos, ptE(0) + chamf_thk, ptE(1) + WinPos + win_clamp_top_max, ptE(0) - 4.5, (2 * ptE(1) + 2 * WinPos + win_clamp_top_max) / 2, pi / 2)
                Call DrawDimLin(ptE(0) + chamf_thk, ptE(1) + WinPos, ptE(0) + chamf_thk, ptE(1) + WinPos - win_clamp_bot_max, ptE(0) - 4.5, (2 * ptE(1) + 2 * WinPos - win_clamp_bot_max) / 2, pi / 2)
            End If
            
    'PLAN VIEWS
    pt13(0) = pt_o(0) + 427 + x - chamf_thk: pt13(1) = pt_o(1) + 42 + ply_thk 'origin point for plan views
    
    'Draw elevation sight [REMOVED]
    'pt33(0) = pt13(0) - x / 2: pt33(1) = pt13(1) + 83.5 - 16
    'Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt33, "VBA_ELEV_C", 1#, 1#, 1#, 0) 'Elevation C for clamps

    'Draw text notes
        'clamps
        pt30(0) = pt13(0) - x: pt30(1) = pt13(1) + y
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES1", 1#, 1#, 1#, 0)
        pt30(0) = pt13(0): pt30(1) = pt13(1)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES2", 1#, 1#, 1#, 0)
        pt30(0) = pt13(0): pt30(1) = pt13(1) + y
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES3", 1#, 1#, 1#, 0)
        pt30(0) = pt13(0) - x: pt30(1) = pt13(1)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES4", 1#, 1#, 1#, 0)
        
        'top clamps
        pt30(0) = pt13(0) - x: pt30(1) = pt13(1) + y + 83.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES5", 1#, 1#, 1#, 0)
        pt30(0) = pt13(0) - (x - 12) / 2: pt30(1) = pt13(1) + y + 83.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES6", 1#, 1#, 1#, 0)
        pt30(0) = pt13(0) - x: pt30(1) = pt13(1) + 83.5
        If SquaringCornerOptionPickingButton.Value = True Then
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES7_LOOP", 1#, 1#, 1#, 0)
        Else
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_PLAN_NOTES7", 1#, 1#, 1#, 0)
        End If
        
    'Draw clamps
    pt15(0) = pt13(0): pt15(1) = pt13(1)
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block_op, 1#, 1#, 1#, 0) 'Insert clamp (operator typical clamps)
    pt15(0) = pt13(0) - x: pt15(1) = pt13(1) + y
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block_bk, 1#, 1#, 1#, 0) 'Insert clamp (back top clamps)
    pt15(0) = pt13(0): pt15(1) = pt13(1) + 83.5
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block_op, 1#, 1#, 1#, 0) 'Insert clamp (operator top clamps)
    If window = True Then Call ChangeProp(BlockRefObj, "Distance1", 2.5 + clamp_L - 12 - y) 'Rotate clamp out if window
    If window = True Then Call ChangeProp(BlockRefObj, "Angle1", swing_ang) 'Rotate clamp out if window
    pt15(0) = pt13(0) - x: pt15(1) = pt13(1) + y + 83.5
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block_bk, 1#, 1#, 1#, 0) 'Insert clamp (back typical clamps)

    'Draw squaring corners and slings for typical clamps
    pt15(0) = pt13(0) - x: pt15(1) = pt13(1)
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, SQ_NAME, 1#, 1#, 1#, 0)
    pt15(0) = pt13(0): pt15(1) = pt13(1) + y
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, SQ_NAME, 1#, 1#, 1#, pi)

    If SquaringCornerOptionRegularButton.Value = True Then
        'Lifting slings for top clamps (regular squaring corners)
        pt15(0) = pt13(0) - x: pt15(1) = pt13(1) + 83.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, SQ_NAME_SLINGS, 1#, 1#, 1#, 0)
        pt15(0) = pt13(0): pt15(1) = pt13(1) + y + 83.5
        If window = False Then
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, SQ_NAME_SLINGS, 1#, 1#, 1#, pi)
        End If
        
        'Add slings to studs for small columns where corner is inversed and sling isn't included at the corner OR for windows
        If Round(x, 3) < 14 Or Round(y, 3) < 14 Or window = True Then
            pt15(0) = pt13(0): pt15(1) = pt13(1) + 83.5 + y
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, "VBA_LIFTING_SLING_C", 1#, 1#, 1#, 0)
            pt15(0) = pt13(0) - x: pt15(1) = pt13(1) + 83.5
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, "VBA_LIFTING_SLING_C", 1#, 1#, 1#, pi)
        End If
    ElseIf SquaringCornerOptionPickingButton.Value = True Then
        'Lifting slings for top clamps (picking loops)
        pt15(0) = pt13(0) - x: pt15(1) = pt13(1) + 83.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, "VBA_PICKING_CORNER", 1#, 1#, 1#, 0)
        pt15(0) = pt13(0): pt15(1) = pt13(1) + y + 83.5
        If window = False Then
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, "VBA_PICKING_CORNER", 1#, 1#, 1#, pi)
        End If
    End If
    'Add sling notes
    If Round(x, 3) >= 14 And Round(y, 3) >= 14 Then
        pt15(0) = pt13(0): pt15(1) = pt13(1) + y + 83.5 'for columns where squaring corners not rotated (normal)
    Else
        pt15(0) = pt13(0) + 6: pt15(1) = pt13(1) - 1 + 83.5 'for columns where squaring corners are rotated, sling is tied between studs
    End If
    If window = False Then
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, "VBA_PLAN_NOTES8", 1#, 1#, 1#, 0)
    End If
    
    'Draw plywood & chamfer
        ' "Clamps" ply & chamfer
        pt12(0) = pt13(0) - x + chamf_thk: pt12(1) = pt13(1) - ply_thk
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, 0) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
        
        pt12(0) = pt12(0) - ply_thk - chamf_thk: pt12(1) = pt12(1) + ply_width_y + ply_thk - chamf_thk - 1.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, -pi / 2) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
        
        pt12(0) = pt12(0) + ply_thk - chamf_thk + ply_width_x - 1.5: pt12(1) = pt12(1) + ply_thk + chamf_thk
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
        
        pt12(0) = pt12(0) + ply_thk + chamf_thk: pt12(1) = pt12(1) - ply_thk + chamf_thk - ply_width_y + 1.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi / 2) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)

        
        ' "Top clamps" ply & chamfer
        pt12(0) = pt12(0) - ply_width_x + 1.5: pt12(1) = pt12(1) + 82
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, 0) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)

        pt12(0) = pt12(0) - ply_thk - chamf_thk: pt12(1) = pt12(1) + ply_width_y - 1.5
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, -pi / 2) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)

        pt12(0) = pt12(0) + ply_width_x - 1.5: pt12(1) = pt12(1) + ply_thk + chamf_thk
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)

        pt12(0) = pt12(0) + ply_thk + chamf_thk: pt12(1) = pt12(1) - ply_width_y + 1.5
        'on error Resume Next
        If window = True Then
            ThisDrawing.ActiveLinetype = ThisDrawing.Linetypes.Item("HIDDEN")
        End If
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, pi / 2) 'Insert ply with chamfer
        Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
        ThisDrawing.ActiveLinetype = ThisDrawing.Linetypes.Item("ByLayer")
        'on error GoTo ErrorHandler
        
    'Draw studs
        ' "Clamps" studs
        pt14(1) = pt13(1) - ply_thk - 1.5
        For i = 1 To n_studs_x - 1 'leave off corner stud on A side
            pt14(0) = pt13(0) - x + stud_spacing_x(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_spax, 1#, 1#, 1#, 0) 'Insert bottom studs
        Next i
        pt14(0) = pt13(0) - x + stud_spacing_x(UBound(stud_spacing_x))
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, 0) 'Insert stud (corner bolt)
            
        pt14(0) = pt13(0) - x - ply_thk - 1.5
        For i = 1 To n_studs_y
            pt14(1) = pt13(1) + y - stud_spacing_y(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_spax, 1#, 1#, 1#, -pi / 2) 'Insert left studs
        Next i
        
        pt14(1) = pt13(1) + y + ply_thk + 1.5
        For i = 1 To n_studs_x
            pt14(0) = pt13(0) - stud_spacing_x(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_spax, 1#, 1#, 1#, pi) 'Insert top studs
        Next i
        
        pt14(0) = pt13(0) + ply_thk + 1.5
        For i = 2 To n_studs_y 'leave off corner stud on B side
            pt14(1) = pt13(1) + stud_spacing_y(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_spax, 1#, 1#, 1#, pi / 2) 'Insert right studs
        Next i
        pt14(1) = pt13(1) + stud_spacing_x(1)
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, pi / 2) 'Insert stud (corner bolt)
        
        ' "Top clamps" studs
        pt13(1) = pt13(1) + 83.5 'move plan view point up for top clamps
        pt14(1) = pt13(1) - ply_thk - 1.5
        For i = 1 To n_studs_x
            pt14(0) = pt13(0) - x + stud_spacing_x(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, 0) 'Insert bottom studs
        Next i
        
        pt14(0) = pt13(0) - x - ply_thk - 1.5
        For i = 1 To n_studs_y
            pt14(1) = pt13(1) + ply_width_y - stud_spacing_y(i) - 1.5
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, -pi / 2) 'Insert left studs
        Next i
        
        pt14(1) = pt13(1) + y + ply_thk + 1.5
        For i = 1 To n_studs_x
            pt14(0) = pt13(0) - stud_spacing_x(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, pi) 'Insert top studs
        Next i
        
        'Switch to hidden style stud blocks
        If window = True Then stud_block_bolt = stud_block_bolt_hidden
        pt14(0) = pt13(0) + ply_thk + 1.5
        For i = 1 To n_studs_y
            pt14(1) = pt13(1) + stud_spacing_y(i)
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block_bolt, 1#, 1#, 1#, pi / 2) 'Insert right studs
        Next i
        'Return to solid stud blocks
        stud_block_bolt = Replace(CStr(stud_block_bolt), "_HIDDEN", "")

        'Draw studs on turned out window clamp (right, y)
        If window = True Then
            For i = 1 To n_studs_y
                pt5(0) = pt13(0) + ply_thk + 1.5 + 1 + Sin(swing_ang) + (3.25 + y - stud_spacing_y(i)) * Cos(swing_ang):  pt5(1) = pt13(1) + y + ply_thk + 2.5 - Cos(swing_ang) + (3.25 + y - stud_spacing_y(i)) * Sin(swing_ang)
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt5, stud_block_bolt, 1#, 1#, 1#, swing_ang + pi) 'Insert right studs
                pt5(0) = pt13(0) + ply_thk + 1.5 + 1 + 2.5 * Sin(swing_ang) + (3.25 + y - chamf_thk) * Cos(swing_ang):   pt5(1) = pt13(1) + y + ply_thk + 2.5 - 2.5 * Cos(swing_ang) + (3.25 + y - chamf_thk) * Sin(swing_ang)
                Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt5, "VBA_PLY_WITH_CHAMFER", 1#, 1#, 1#, swing_ang + pi) 'Insert plywood
                Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
            Next i
        End If
        
    'Draw brace clamps
    
    'Draw brace clamps in "CLAMPS" plan view if attached to clamp, otherwise stay at top clamps
    If brace_clamp > n_top_clamps Then
        pt13(1) = pt13(1) - 83.5
    End If
    
    'The 3 clamp locations are: (1) upper left stud on the x side
    '                           (2) upper right stud OR clamp tail on the x side
    '                           (3) lower left stud on the y side
    
    If clamp_L - 12.5 - x >= 8.5 Then 'if clamp "tail" is long enough, place this brace on tail
        pt16(0) = pt13(0) - x - 6.25 + clamp_L - 4.25: pt16(1) = pt13(1) + ply_thk + y
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_AFB_PLAN", 1#, 1#, 1#, 0)
    Else
        If Round(stud_spacing_x(2) - stud_spacing_x(1) - 3.5, 3) >= min_stud_gap Then 'Otherwise check if there is room to place AFB by upper right stud
            pt16(0) = pt13(0) - stud_spacing_x(1) - 3.5 - (stud_spacing_x(2) - stud_spacing_x(1) - 3.5) / 2: pt16(1) = pt13(1) + ply_thk + y
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_AFB_PLAN", 1#, 1#, 1#, 0)
        Else 'If no room on far stud, force it onto the clamp tail anyways with some offset
            pt16(0) = pt13(0) - x - 6.25 + clamp_L - 2: pt16(1) = pt13(1) + ply_thk + y
            Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_AFB_PLAN", 1#, 1#, 1#, 0)
        End If
    End If
    
    pt16(0) = pt13(0) - x + 3.8125:     pt16(1) = pt13(1) + ply_thk + y
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_AFB_PLAN", 1#, 1#, 1#, 0) 'Insert AFB by upper left x stud

    pt16(0) = pt13(0) - x - ply_thk:    pt16(1) = pt13(1) + 3.8125
    Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_AFB_PLAN", 1#, 1#, 1#, pi / 2) 'Insert AFB by lower left y stud

    'Add plan note 9 for AFB
    pt16(0) = pt13(0) - x: pt16(1) = pt13(1)
    If y > 27 Then
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_PLAN_NOTES9A", 1#, 1#, 1#, 0)
    Else
        pt16(1) = pt13(1) + y
        Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt16, "VBA_PLAN_NOTES9B", 1#, 1#, 1#, 0)
    End If
    
    'Fix location for the rest of plan view drawing if the AFBs were drawn in the "TOP CLAMPS" plan view
    If brace_clamp <= n_top_clamps Then
        pt13(1) = pt13(1) - 83.5
    End If

    ThisDrawing.ActiveLayer = ThisDrawing.Layers("CONC")
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
        MTextObject2.StyleName = "Arial Narrow Bold": MTextObject2.Height = 1.5: MTextObject2.AttachmentPoint = acAttachmentPointTopCenter
    
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
Public Function ChangeProp(BlockRefObj, PropName, PropValue)

    Props = BlockRefObj.GetDynamicBlockProperties
    Dim oprop As AcadDynamicBlockReferenceProperty
    
    For i = LBound(Props) To UBound(Props)
        Set oprop = Props(i) 'Get object properties from block one at a time
        If oprop.PropertyName = PropName Then 'compare to property name
            oprop.Value = PropValue 'if name matches, assign new PropValue
        End If
    'Exit For
    Next i
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
'draws a rectangle with lower left corner at xoff, yoff with COORDINATES for far point specified
'    4_______3B
'    |       |
'    |       |
'    |_______|
'   A1       2

    Public Function DrawRectangle(xA As Double, yA As Double, xB As Double, yB As Double)
        Dim plineObj As AcadLWPolyline
        Dim pts(0 To 7) As Double
        
        pts(0) = xA#:   pts(1) = yA
        pts(2) = xB#:   pts(3) = yA
        pts(4) = xB#:   pts(5) = yB
        pts(6) = xA#:   pts(7) = yB

        Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
        plineObj.Closed = True

    End Function
Public Function FindMax(Arr As Variant)
    Dim RetMax As Double
    RetMax = Arr(LBound(Arr)) 'initialize with first value
    
    For i = LBound(Arr) To UBound(Arr)
    If Arr(i) > RetMax Then
        RetMax = Arr(i)
    End If
    Next i
    FindMax = RetMax
End Function
Public Function FindMin(Arr As Variant)
    Dim RetMin As Double
    RetMin = Arr(LBound(Arr)) 'initialize with first value
    For i = LBound(Arr) To UBound(Arr)
    If Arr(i) < RetMin Then
        RetMin = Arr(i)
    End If
    Next i
    FindMin = RetMin
End Function
Public Function RoundDown(num As Double) As Integer
    num = Int(num)
    RoundDown = num
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