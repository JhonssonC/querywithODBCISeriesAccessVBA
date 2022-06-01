Attribute VB_Name = "SOL"

Private Function Col_Letter(lngCol As Variant)
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function



Private Function QUERY(sql)


    
    Dim conn As Object 'Variable for ADODB.Connection object
    Dim RS As Object 'Variable for ADODB.Recordset object
    Dim USERNAME, PASSWORD
    
    Dim arr(), arr2()
    arr = Array()
    
    Set conn = CreateObject("ADODB.Connection")
    Set RS = CreateObject("ADODB.Recordset")
    
    FILA = ActiveCell.Row
    
    USERNAME = Sheets("VAR").Range("E1")
    PASSWORD = Sheets("VAR").Range("E2")
    
    conn.ConnectionString = "dsn=SICOPUB-MAN;User Id=" & USERNAME & ";Password=" & PASSWORD & " ;"
    
    
    On Error GoTo ERROR1
    
    conn.Open
    
    
    RS.Open sql, conn
    
    Dim I, j, K, l
    
    I = 0
    j = 0
    
    
    
    With RS
      If Not (.EOF And .BOF) Then
            arr2 = .GetRows
          
            K = UBound(arr2, 2)
            l = UBound(arr2, 1)
            
            ReDim arr(K, l)
            
            For I = 0 To K
                For j = 0 To l
                    arr(I, j) = arr2(j, I)
                Next
            Next
          
      Else
          arr = Array()
      End If
      .Close
    End With
    
    

    
    QUERY = arr
    
    'RS.Close
    conn.Close
    
ERROR1:

     QUERY = arr
    
End Function



Sub LASTSOL()

    
    If Selection.Cells.Rows.Count > 1 Then
        
        RF = Selection.Cells(Selection.Cells.Rows.Count, 1).Row
        RI = Selection.Cells(1, 1).Row
        
        COLU = Col_Letter(Selection.Cells.Column)
        
        Dim celda As Range
    
        For Each celda In Range(COLU & RI & ":" & COLU & RF).SpecialCells(xlCellTypeVisible)
            
            celda.Select
            FILA = ActiveCell.Row
            Range(Sheets("VAR").Range("B2") & FILA).Select
            
            If ActiveCell <> "" Then
            
                CLIENT = ActiveCell
                
                If sql = "" Then
                
                    sql = "" & Sheets("VAR").Range("B1")
                    sql = Replace(sql, "[[CODE]]", CLIENT)
                    
                End If
                
                For F = 0 To 10
                    RS = QUERY(sql)
                    If Not IsEmpty(RS) Then
                        If UBound(RS) > -1 Then
                            Exit For
                        End If
                    End If
                Next
                
                If Not (IsEmpty(RS)) Then
                
                 
                    Range(Sheets("VAR").Range("B3") & ActiveCell.Row) = RS(0, 0)
                    Range(Sheets("VAR").Range("B4") & ActiveCell.Row) = RS(0, 1)
                    Range(Sheets("VAR").Range("B5") & ActiveCell.Row) = RS(0, 2)
                    Range(Sheets("VAR").Range("B6") & ActiveCell.Row) = RS(0, 3)
                    Range(Sheets("VAR").Range("B7") & ActiveCell.Row) = RS(0, 4)
                    
                    
                End If
                
                
                sql = ""
                
                
            End If
            
        Next
        
        
    ElseIf ActiveCell <> "" Then
                

        CLIENT = ActiveCell
                
        If sql = "" Then
        
            sql = "" & Sheets("VAR").Range("B1")
            sql = Replace(sql, "[[CODE]]", CLIENT)
            
        End If
        
        For F = 0 To 10
            RS = QUERY(sql)
            If Not IsEmpty(RS) Then
                If UBound(RS) > -1 Then
                    Exit For
                End If
            End If
        Next
        
        If Not (IsEmpty(RS)) Then
        
            
            Range(Sheets("VAR").Range("B3") & ActiveCell.Row) = RS(0, 0)
            Range(Sheets("VAR").Range("B4") & ActiveCell.Row) = RS(0, 1)
            Range(Sheets("VAR").Range("B5") & ActiveCell.Row) = RS(0, 2)
            Range(Sheets("VAR").Range("B6") & ActiveCell.Row) = RS(0, 3)
            Range(Sheets("VAR").Range("B7") & ActiveCell.Row) = RS(0, 4)

            
        End If
        
        
        sql = ""
        
     
    End If
End Sub



