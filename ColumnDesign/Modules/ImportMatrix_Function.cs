/*namespace ColumnDesign.Modules
{
    public class ImportMatrix_Function
    {
        
Public Function ImportMatrix(filepath) 'filepath string is specified in the main routine
    'Import a .csv file
    'Ignore first column and first row
    'All rows and columns must have the same number of entries
    'All values must be numeric
    
    Dim i As Integer: i = 0 'number of rows start at minimum)
    Dim j As Integer        'number of columns (start at minimum)
    Dim n_rows_tot As Integer: n_rows_tot = 0
    Dim a As Integer
    Dim b As Integer
    
    Dim mystring As String
    Dim temp_row() As String
    Dim return_matrix() As Integer

    'Find the total number of rows and columns to redim the array to the proper size
    Open filepath For Input As #1
    Do While Not EOF(1)
        Line Input #1, mystring
        n_rows_tot = n_rows_tot + 1
    Loop
    Close #1
    temp_row = Split(mystring, ",") 'Load 1 row to read its length
    
    ReDim return_matrix(n_rows_tot - 1, UBound(temp_row)) 'Redim array to maximum file rows and columns (n_rows_tot and UBound(temp_row) respectively)
                                                          'Can only modify LAST dimension using redim with preserve.
                                                          'This avoids constantly redimming the array with preserve and worrying about which dimension is being changed.
    
    Open filepath For Input As #1           'Reopens the file to start a new loop.
    Do While Not EOF(1)                     'Do while the end of file hasn't been reached
        Line Input #1, mystring             'Reads file one full row at a time as a string
        temp_row = Split(mystring, ",")     'Deliminates the row by "," and stores each value into the temp_row array as a string.

        For j = 0 To UBound(temp_row)                               'Cycle through each entry in a single row
            return_matrix(i, j) = CInt(temp_row(j))                 'Place each entry of temp_row into the output matrix and convert it from string to integer
        Next j

        i = i + 1 'Count ticks up for the next matrix row
    Loop
    
    ImportMatrix = return_matrix 'Reassigns matrix to function name for output
    
    Close #1 'Close the file
    
End Function


    }
}*/