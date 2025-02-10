Sub UnisciFogli()
    Dim ws As Worksheet
    Dim wsUnico As Worksheet
    Dim rng As Range
    Dim ultimaRiga As Long
    
    ' Crea un nuovo foglio di lavoro per unire i dati
    Set wsUnico = ThisWorkbook.Worksheets.Add
    wsUnico.Name = "Unione"
    
    ' Imposta la riga iniziale per l'incollaggio dei dati
    ultimaRiga = 1
    
    ' Loop attraverso tutti i fogli di lavoro
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsUnico.Name Then
            ' Trova l'ultima riga con dati nel foglio corrente
            Set rng = ws.UsedRange
            rng.Copy wsUnico.Cells(ultimaRiga, 1)
            
            ' Aggiorna l'ultima riga per il prossimo incollaggio
            ultimaRiga = ultimaRiga + rng.Rows.Count
        End If
    Next ws
    
    MsgBox "Unione completata!"
End Sub
