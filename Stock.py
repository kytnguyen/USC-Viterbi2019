Sub Stock()

' Set an initial variable for holding the ticket name

Dim Ticket_Name As String

' Set an initial variable for holding the total stock volume

Dim Volume As Double
Volume = 0

' Keep track of the location for each stock ticket in the summary table

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Count the rows

Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all credit card purchases

For i = 2 To last_row

' Check if we are still within the same ticket, if it is not...

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the Brand name

        Ticket_Name = Cells(i, 1).Value

        ' Add to the Volume

        Volume = Volume + Cells(i, 7).Value

        ' Print the Credit Card Brand in the Summary Table

        Range("I" & Summary_Table_Row).Value = Ticket_Name
        
        ' Print the Brand Amount to the Summary Table
        
        Range("J" & Summary_Table_Row).Value = Volume

        ' Add one to the summary table row

        Summary_Table_Row = Summary_Table_Row + 1
            
        ' Reset the Brand Total

        Volume = 0

' If the cell immediately following a row is the same ticket...

        Else

        ' Add to the Brand Total

        Volume = Volume + Cells(i, 7).Value

        End If

Next i

End Sub


