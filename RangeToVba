Sub ParseRangeToJson()

'OutPut variable
Dim Json As String
Dim cell As Range

Json = "[" & Chr(10)

'Count amount of used rows by given column as reference
For I = 1 To Application.Range("A1:A" & Range("A1").End(xlDown).Row).Count - 1

    Json = Json & "{" & Chr(10)
        
        'Range of columns to parse
        For Each cell In Range("A1:G1")
            
            'Get columns name
            Json = Json & Chr(34) & cell.Value & Chr(34) & ": "
            'Get data to columns
            Json = Json & Chr(34) & cell.Offset(I, 0).Value & Chr(34) & "," & Chr(10)

        Next cell

    Json = Left(Json, Len(Json) - 2)
    Json = Json & Chr(10) & "}," & Chr(10)

Next I
 
Json = Left(Json, Len(Json) - 2)
Json = Json & Chr(10) & "]"

'Json output
Debug.Print Json

End Sub
