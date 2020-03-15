Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         max_prof = 0
         min_prof = 0
         max_vol = 0
         max_tick = ""
         min_tick = ""
         max_vol_tick = ""
         For I = 1 To WS_Count
            newTicker = ""
            Dim sumTotal As Long
            counter = 1
            percentage_change = 0
            endTotal = 0
            setChange = 0
            ' MsgBox (Rows.Count)
            For J = 2 To Rows.Count
                ' MsgBox (Cells(J, 2).Value)
                newVal = Cells(J, 1).Value
                If newTicker <> newVal Then
                    If J = 2 Then
                        prevCost = Cells(J, 3).Value
                        startPrice = Cells(J, 3).Value
                        newTicker = Cells(J, 1).Value
                        ' MsgBox (newTicker)
                        
                    Else
                        oldStartPrice = startPrice
                        startPrice = Cells(J, 3).Value
                        endTotal_prev = Cells(J - 1, 6).Value
                        setChange = 1
                        
                    End If
                    
                    Cells(counter, 8).Value = newTicker
                    
                    ' MsgBox (Cells(J, 1).Value)
                    newTicker = Cells(J, 1).Value
                    counter = counter + 1
                    Cells(counter - 1, 11).Value = sumTotal
                    If sumTotal > max_vol Then
                        max_vol = sumTotal
                        max_vol_tick = Cells(counter - 1, 1).Value
                    End If
                    
                    sumTotal = Cells(J, 7).Value
                    If setChange = 1 Then
                        profit = endTotal_prev - oldStartPrice
                         Dim myVal As Double
                         myVal = profit / oldStartPrice * 100
                         If myVal < 0 Then
                            Cells(counter - 1, 9).Interior.Color = vbRed
                         Else
                          Cells(counter - 1, 9).Interior.Color = vbGreen
                         End If
                     Cells(counter, 8) = newTicker
                     Cells(counter - 1, 9) = profit
                     Cells(counter - 1, 10) = myVal
                     If myVal > max_prof Then
                        max_tick = Cells(counter - 1, 8).Value
                        max_prof = myVal
                     End If
                     If myVal < min_prof Then
                        min_val = myVal
                        min_tick = Cells(counter - 1, 8).Value
                     End If
                        
                     
                     
                     setChange = 0
                     End If
                Else
                myCell = Cells(J, 7).Value
                   sumTotal = sumTotal + myCell
                   If sumTotal > max_vol Then
                    max_vol = sum_total
                    
                    End If
                End If
            Next J
         Next I


        Cells(2, 13).Value = "Max Profit "
        Cells(3, 13).Value = "Min Profit "
        Cells(4, 13).Value = "Max Vol "
        Cells(2, 14).Value = max_tick
        Cells(3, 14).Value = min_tick
        Cells(4, 14).Value = max_vol_tick
        Cells(2, 15).Value = max_prof
        Cells(3, 15).Value = min_prof
        Cells(4, 15).Value = max_vol
        
End Sub