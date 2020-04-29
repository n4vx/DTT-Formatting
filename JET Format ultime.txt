Sub formattingGL()


i = 0

Cells(1, 1).Select

Columns(ActiveCell.Column).EntireColumn.Delete

Do While ActiveCell.Value <> "Réf."

        Rows(ActiveCell.Row).EntireRow.Delete

                Loop
                

ActiveCell.End(xlDown).Select



Do While ActiveCell.Offset(0, 2).End(xlDown) <> ""


    If ActiveCell.Value = "Pièce" Or ActiveCell.Value = "Réf." Or IsEmpty(ActiveCell) Then
    
    Rows(ActiveCell.Row).EntireRow.Delete
    
      Else
        
        a = ActiveCell.Offset(-1, 0).Value
        b = ActiveCell.Offset(-1, 2).Value
        c = ActiveCell.Offset(-1, 1).Value

'mettre un not avant isnumeric pour les numeros de compte

        If Not IsNumeric(ActiveCell.Value) Then
    
            ActiveCell.Value = a
            ActiveCell.Offset(0, 1).Value = c
            
        Else
                
                If IsNumeric(ActiveCell.Offset(0, 1).Value) Then
                
                        ActiveCell.Offset(0, 1).Value = ActiveCell.Offset(0, 2).Value
                        End If
                        
              ActiveCell.Offset(1, 0).Select
    
    End If
     End If
 

Loop

ActiveCell.End(xlUp).Select
ActiveCell.End(xlUp).Select


Do While ActiveCell.End(xlDown) <> ""

    If IsEmpty(ActiveCell.Offset(0, 3)) Then
        Rows(ActiveCell.Row).EntireRow.Delete
        Else
        ActiveCell.Offset(1, 0).Select
        End If
        
        Loop
        
ActiveCell.Offset(1, 0).Select
Rows(ActiveCell.Row).EntireRow.Delete

ActiveCell.End(xlUp).Select
ActiveCell.End(xlUp).Select

ActiveCell.Offset(-1, 0).Select
Rows(ActiveCell.Row).EntireRow.Delete

ActiveCell.Offset(-1, 0).Value = "Acc number"
ActiveCell.Offset(-1, 1).Value = "Acc name"
ActiveCell.Offset(-1, 2).Value = "Date"

Cells(1, 1).Select

Do While ActiveCell.End(xlToRight) <> ""
    If Not ActiveCell.End(xlDown) <> "" Then
    
     Columns(ActiveCell.Column).EntireColumn.Delete
     Else
     
     ActiveCell.Offset(0, 1).Select
    End If
    Loop
    
    


End Sub


