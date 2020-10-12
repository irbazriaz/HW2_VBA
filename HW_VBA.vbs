Sub HW()


    For Each ws In Worksheets
        Dim i As Long
            Dim j As Integer


    ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
                ws.Range("L1").Value = "Total"
                    ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"


    Dim ticker As String
        Dim total As Double
            total = 0

             
    Dim openprice As Double
        openprice = ws.Cells(2, 3).Value
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double


Dim incrticker As String
    Dim decrticker As String
        Dim incrvolticker As String
            'Dim decrvolticker As String
            Dim incrpercent As Double
        Dim decrpercent As Double
    Dim incrvol As Double
'Dim decrvol As Double

incrpercent = 0
decrpercent = 0
incrvol = 0

    Dim row As Integer
        row = 2
  
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
             ticker = ws.Cells(i, 1).Value
                total = total + ws.Cells(i, 7).Value
                    ws.Range("I" & row).Value = ticker
                        ws.Range("L" & row).Value = total
                    closeprice = ws.Cells(i, 6).Value
                yearlychange = closeprice - openprice
            ws.Range("J" & row).Value = yearlychange

        If yearlychange > 0 Then
                
            ws.Range("J" & row).Interior.ColorIndex = 4

        ElseIf yearlychange < 0 Then
                
                ws.Range("J" & row).Interior.ColorIndex = 3

        End If


        If openprice = 0 Then
                
            percentchange = 0
                
        Else
            
            percentchange = yearlychange / openprice
                
        End If

ws.Range("K" & row).Value = percentchange
ws.Range("K" & row).NumberFormat = "0.00%"
   
              
row = row + 1
total = 0
openprice = ws.Cells(i + 1, 3)
            
        Else
              
           total = total + ws.Cells(i, 7).Value
 
        End If
        
        If percentchange > incrpercent Then

incrpercent = percentchange
    incrticker = ticker
        ws.Range("P2").Value = incrticker
            ws.Range("Q2").Value = incrpercent
        ws.Range("Q2").Interior.ColorIndex = 4
    ws.Range("Q2").NumberFormat = "0.00%"

    ElseIf percentchange < decrpercent Then

    decrpercent = percentchange
        decrticker = ticker
            ws.Range("P3").Value = decrticker
        ws.Range("Q3").Value = decrpercent
    ws.Range("Q3").Interior.ColorIndex = 3
ws.Range("Q3").NumberFormat = "0.00%"

    End If

    If total > incrvol Then

incrvol = total
    incrvolticker = ticker
    ws.Range("P4").Value = incrvolticker
ws.Range("Q4").Value = incrvol


            End If
        
        Next i
    
    Next ws
        
End Sub