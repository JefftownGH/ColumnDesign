/*namespace ColumnDesign.Modules
{
    public class ConvertFeetInchesToNumber
    {
        
Function ConvertToNum(StringInput) As Double

    '####################################################################################################################################################################################
    'This function has the added ability to consider fractions e.g. 10'-5 1/2"
    'str_temp1 = 10
    'str_temp2 = 5
    'str_temp3 = 1
    'str_temp4 = 2
    '####################################################################################################################################################################################
    
    'On Error GoTo ConvertToNumErrorHandler
    
    Dim neg_tog As Integer
    Dim found_num As Integer
    Dim str_temp1 As String
    Dim str_temp2 As String
    Dim str_temp3 As String
    Dim str_temp4 As String

    neg_tog = 1
    
    'Convert empty values or just a . to 0
    If StringInput = "" Or StringInput = " " Or StringInput = "." Then
        ConvertToNum = 0
        Exit Function
    End If
    
    var_in = StringInput 'Input value. This function can handle the following formats:
        '125                125
        '125.5              125.5
        '125.5"             125.5
        '10'                120
        '10'-5.5"           125.5
        '10'5.5             125.5
        '10'5.5"            125.5
        '10'-5.5            125.5
        '10 5.5             125.5
        '10'.5              120.5
        '10'-5 1/2"         125.5
        '10'5 1/2"          125.5
        '10'5-1/2"          125.5
        '10' 1/2"           120.5
        '10' 5.5/2"         122.75
        '10' 5.5 5/2"       128
        '10' 5.5 5.5/2"     128.25
        
    str_temp1 = "" 'Reset to blank strings
    str_temp2 = ""
    str_temp3 = ""
    str_temp4 = ""
    
    'Check if string contains no numeric characters, then the output is 0
    For iCnt0 = 1 To Len(var_in)
        If IsNumeric(Mid(var_in, iCnt0, 1)) Then
            GoTo TheresANumberInThere
        ElseIf iCnt0 = Len(var_in) Then
            out_num = 0
            GoTo SkipConvertToNumErrorHandler
        End If
    Next iCnt0
TheresANumberInThere:

    'Check if number is negative, if it is, remove the negative sign and process the positive number, then make it negative again later
    If Mid(var_in, 1, 1) = "-" Then
        neg_tog = -1
        var_in = Right(var_in, Len(var_in) - 1)
    End If

    'Read the input box 1 character at a time left to right. If the char is numeric or a ".", add it to str_temp1. If not, stop reading
    found_num = 0
    For iCnt1 = 1 To Len(var_in)
        If IsNumeric(Mid(var_in, iCnt1, 1)) Or Mid(var_in, iCnt1, 1) = "." Then 'Record numbers and periods
            str_temp1 = str_temp1 & Mid(var_in, iCnt1, 1)
            found_num = 1
        ElseIf found_num = 0 Then 'If no number has been found yet, do nothing, regardless of character
        ElseIf found_num = 1 Then 'If numbers have been found, and the character isn't a number of period, exit the loop
            Exit For
        End If
    Next iCnt1
    
    'Start again 1 char after where the last loop left off and store numeric characters until a non-numeric, non-period character is encountered
    found_num = 0
    For iCnt2 = iCnt1 + 1 To Len(var_in)
        If IsNumeric(Mid(var_in, iCnt2, 1)) Or Mid(var_in, iCnt2, 1) = "." Then 'Record numbers and periods
            str_temp2 = str_temp2 & Mid(var_in, iCnt2, 1)
            found_num = 1
        ElseIf found_num = 0 Then 'If no number has been found yet, do nothing, regardless of character
        ElseIf found_num = 1 Then 'If numbers have been found, and the character isn't a number of period, exit the loop
            Exit For
        End If
    Next iCnt2
    
    'str_temp3 and str_temp4 are for fractions and may not be used
    'Start again 1 char after where the last loop left off
    found_num = 0
    For iCnt3 = iCnt2 + 1 To Len(var_in)
        If IsNumeric(Mid(var_in, iCnt3, 1)) Or Mid(var_in, iCnt3, 1) = "." Then 'Record numbers and periods
            str_temp3 = str_temp3 & Mid(var_in, iCnt3, 1)
            found_num = 1
        ElseIf found_num = 0 Then 'If no number has been found yet, do nothing, regardless of character
        ElseIf found_num = 1 Then 'If numbers have been found, and the character isn't a number of period, exit the loop
            Exit For
        End If
    Next iCnt3
    
    found_num = 0
    For iCnt4 = iCnt3 + 1 To Len(var_in)
        If IsNumeric(Mid(var_in, iCnt4, 1)) Or Mid(var_in, iCnt4, 1) = "." Then 'Record numbers and periods
            str_temp4 = str_temp4 & Mid(var_in, iCnt4, 1)
            found_num = 1
        ElseIf found_num = 0 Then 'If no number has been found yet, do nothing, regardless of character
        ElseIf found_num = 1 Then 'If numbers have been found, and the character isn't a number of period, exit the loop
            Exit For
        End If
    Next iCnt4
    
    'MsgBox "str_temp3: " & str_temp3 & vbNewLine & "str_temp4: " & str_temp4
    
    If str_temp1 <> "" And str_temp1 <> "." And _
       str_temp2 <> "" And str_temp2 <> "." And _
       str_temp3 <> "" And str_temp3 <> "." And _
       str_temp4 <> "" And str_temp4 <> "." Then
            If CDbl(str_temp4) = 0 Then
                out_num = 0
            Else
                out_num = 12 * CDbl(str_temp1) + CDbl(str_temp2) + CDbl(str_temp3) / CDbl(str_temp4)
            End If
            
    ElseIf str_temp1 <> "" And str_temp1 <> "." And _
           str_temp2 <> "" And str_temp2 <> "." And _
           str_temp3 <> "" And str_temp3 <> "." Then
            If CDbl(str_temp3) = 0 Then
                out_num = 0
            Else
                out_num = 12 * CDbl(str_temp1) + CDbl(str_temp2) / CDbl(str_temp3)
            End If

    ElseIf str_temp1 <> "" And str_temp1 <> "." And _
           str_temp2 <> "" And str_temp2 <> "." Then
            out_num = 12 * CDbl(str_temp1) + CDbl(str_temp2)
            
    ElseIf str_temp1 <> "" And str_temp1 <> "." Then
        If Mid(var_in, iCnt1, 1) = "'" Then
            out_num = CDbl(str_temp1) * 12
        Else
            out_num = CDbl(str_temp1)
        End If
    Else
        out_num = 0
    End If
  
    'Apply the negative value and set the function output
    ConvertToNum = neg_tog * out_num
    
    GoTo SkipConvertToNumErrorHandler
ConvertToNumErrorHandler:
        Exit Function
    On Error GoTo 0
SkipConvertToNumErrorHandler:
End Function
    }
}*/