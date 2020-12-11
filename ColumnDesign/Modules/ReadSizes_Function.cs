/*namespace ColumnDesign.Modules
{
    public class ReadSizes_Function
    {
        
        Public Function ReadSizes(CSV_Ply) As Double() 'Initialize output array (Each line is 1 plywood size.
            Dim TempStr1 As String
            Dim TempStr2() As String
        Dim TempNum1() As Double
        Dim TempOutput() As Double

        On Error GoTo AbortReading
    
        'Abort if string is empty
            If CSV_Ply = "" Then
            AbortReading:
        ReDim ReadSizes(1 To 1)
        ReDim TempOutput(1 To 1)
        TempOutput(1) = 0
        ReadSizes = TempOutput
            Exit Function
            End If
    
        'Read the given ply spacing and insert it into the ply_spacing array
        j = 1
        For i = 1 To Len(CSV_Ply)
            'Mid reads a specified number of characters from a string given a starting point.
        'In this case, mid starts reading at i and only reads 1 character at a time.
        'On commas or the last character, commit everything read up to that point
        If Mid(CSV_Ply, i, 1) = "," Or i = Len(CSV_Ply) Then
            If i = Len(CSV_Ply) Then
            TempStr1 = TempStr1 & Mid(CSV_Ply, i, 1)
        End If
        ReDim Preserve TempStr2(1 To j)
        TempStr2(j) = TempStr1
            j = j + 1
        TempStr1 = ""
        Else
            TempStr1 = TempStr1 & Mid(CSV_Ply, i, 1)
        End If
        Next i
    
        'Check if too many plywood dims are used
        If UBound(TempStr2) > 11 Then
            MsgBox ("Error: Too many ply sheets defined (max: 11).")
        GoTo AbortReading
        End If
    
        'Fill output array
            ReDim TempOutput(1 To UBound(TempStr2))
        ReDim ReadSizes(1 To UBound(TempStr2))
        For i = 1 To UBound(TempStr2)
        TempOutput(i) = ConvertToNum(TempStr2(i))
        Next i
    
        'Assign output variable
            ReadSizes = TempOutput
        
        On Error GoTo 0
        End Function

    }
}*/