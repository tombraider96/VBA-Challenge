Sub Challenge2Hard()


For Each ws In Worksheets
        Dim WorksheetName As String
        WorksheetName = ws.Name
        
        Sheets(ws.Name).Select


Dim DateMinOpen As Variant
Dim DateMaxClose As Variant
Dim i As Double


Dim x As Double


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Volume"


x = 2
i = 2

Cells(x, 9).Value = Cells(x, 1).Value

DateMinOpen = Cells(i, 3).Value



LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow





If Cells(i, 1).Value = Cells(x, 9).Value Then


TotalV = TotalV + Cells(i, 7).Value


DateMaxClose = Cells(i, 6).Value


     Else
     


Cells(x, 10).Value = DateMaxClose - DateMinOpen

                If DateMaxClose <= 0 Then
            
                    Cells(x, 11).Value = 0
                    
                    Else
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                End If
                
                    Cells(x, 11).Style = "Percent"
                        
            If Cells(x, 10).Value >= 0 Then
                                
                Cells(x, 10).Interior.ColorIndex = 4
                                    
                    Else
                                
                Cells(x, 10).Interior.ColorIndex = 3
                    
            End If
                
Cells(x, 12).Value = TotalV


DateMinOpen = Cells(i, 3).Value

TotalV = Cells(i, 7).Value

x = x + 1
Cells(x, 9).Value = Cells(i, 1).Value

End If

Next i

Cells(x, 10).Value = DateMaxClose - DateMinOpen

                If DateMaxClose <= 0 Then
            
                    Cells(x, 11).Value = 0
                    
                    Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Cells(x, 11).Value = (DateMaxClose / DateMinOpen) - 1
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                End If
                
                    Cells(x, 11).Style = "Percent"
                        
            If Cells(x, 10).Value >= 0 Then
                                
                Cells(x, 10).Interior.ColorIndex = 4
                                    
                    Else
                                
                Cells(x, 10).Interior.ColorIndex = 3
                    
            End If
                
Cells(x, 12).Value = TotalV

   Volume_Greatest_Decrease = 100000
        Ticker_Greatest_Decrease = 100000
        
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        
        For x = 2 To LastRow
        
        
        If Cells(x, 11).Value > Volume_Greatest_Increase Then
            
            Ticker_Greatest_Increase = Cells(x, 9).Value
            Volume_Greatest_Increase = Cells(x, 11).Value
        
        End If
        
        
        If Cells(x, 11).Value < Volume_Greatest_Decrease Then
            
            Ticker_Greatest_Decrease = Cells(x, 9).Value
            Volume_Greatest_Decrease = Cells(x, 11).Value
        
        End If
        
        
        If Cells(x, 12).Value > Volume_Greatest_Total_Volume Then
            
            Ticker_Greatest_Total_Volume = Cells(x, 9).Value
            Volume_Greatest_Total_Volume = Cells(x, 12).Value
        
        End If
        
        Next x
        
Cells(2, 16).Value = Ticker_Greatest_Increase
Cells(2, 17).Value = Volume_Greatest_Increase
Cells(2, 17).Style = "Percent"
Cells(3, 16).Value = Ticker_Greatest_Decrease
Cells(3, 17).Value = Volume_Greatest_Decrease
Cells(3, 17).Style = "Percent"
Cells(4, 16).Value = Ticker_Greatest_Total_Volume
Cells(4, 17).Value = Volume_Greatest_Total_Volume

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next ws
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub