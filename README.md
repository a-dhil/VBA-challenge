# VBA-challenge
Here is the link to my repo: https://github.com/a-dhil/VBA-challenge
name of the file with code is: Challenge_2_Updated_macro.bas. (This is a bas extension file with code of only macro. I was not able to upload excel file as it was too large. Teaching Assitant told me that this will work.)
I finised working on this assignment with the help of tutor, zoom class recording, ChatGPT and AskBCS Learning Assistance.
 I had an appointment with a tutor on Friday who help me understand the assignment in detail with all the specific requirements and an overview of git.
 Zoom recording of the class help me get start writing the code. 
 Once the logic was clear and i was able to write the beginning of the code and dispay each ticker symbol in summary table. With the help of in class activities especially the credit card activity and last activity of US census.
 I calculated all the formulas and had an idea how to use them. Got stuck in how to display the values of first day opening price and last day closing price after calculating them.
 Then I added the part of my code on ChatGPT and asked it for help. 
 There was a lot of trail and error while doing this.
 I was working on a excel sheet with only three stocks with three days price, this was a very small data set and i was able to manuplate it quicky and was helpful in asking ChatGPT what exactly I was asking for.
 After my code was running on with this small data I copied the code and pasted it on alphabetical_testing file. It wroked on it. Then I proceded onto copy code from here and past it on orginal excel sheet.
 In the end I got stuck at summary table finding greatest precent increase and decrease. Where ASKBCS assistant helped me writing the last part of code.
 About how to use min and max functions and how to get index of the min and max values and display them.
Here is the some example of ChatGPT instructions(they are not in ther order i asked)
"as you can see there is an empty column name yearly change i want to use a formula and put yearly change there.  The formula for that is(last day closing price of the year - first day opening price). I will give you the code of what i have done so far. Do say anything just understand problem first"
"this code is not working the way i want it to be it is printing ticker symbol in summary table which is correct but it is not calculating yearly change as i told you to. For that  closing price of the last day minus opening price on day one.""
"this code is still wrong i want  -0.04 in row 2 of summary table in cell yearly chnage for AAB. and second row of summary table in front of AAF i want its yearly chnage which is -0.04 too. I got this by last day close price minus first day open price"
I was working with very small sample data so it was easy to put it on chapGPT and ask it to mauplate the way i want.

here is my code
Sub sample_stock():

    For Each ws In Worksheets

       
        Dim WorksheetName As String
        Dim LastRow As Double
       
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        WorksheetName = ws.Name
        
    
      
        ws.Cells(1, 9).Value = "Ticker"
        
        ws.Cells(1, 10).Value = "Yearly Change"
        
        ws.Cells(1, 11).Value = "Percent Change"
        
       
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim ticker_symbol As String
        Dim yearly_change As Double
        Dim summary_table_row As Integer
        Dim opening_price As Double
        Dim closing_price As Double
        Dim total_volume As Double
        
    
        
        summary_table_row = 2
        
        
        
        
        For i = 2 To LastRow
            
            
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                If i > 2 Then
                    ws.Cells(summary_table_row - 1, 12).Value = total_volume
                End If
                    
            
                ticker_symbol = ws.Cells(i, 1).Value
                opening_price = ws.Cells(i, 3).Value
                total_volume = 0
                
            End If
                total_volume = total_volume + CDbl(ws.Cells(i, 7).Value)
            
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                closing_price = ws.Cells(i, 6).Value
                yearly_change = closing_price - opening_price
                
            
                ws.Cells(summary_table_row, 9).Value = ticker_symbol
                ws.Cells(summary_table_row, 10).Value = yearly_change
                
                 If yearly_change < 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3  ' Red
                ElseIf yearly_change > 0 Then
                    ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4 'Green
                End If
                
                If opening_price <> 0 Then
                    ws.Cells(summary_table_row, 11).Value = Round((yearly_change / opening_price) * 100, 2) & "%"
                Else
                    ws.Cells(summary_table_row, 11).Value = "0.00%"
                End If
                
                
                
                summary_table_row = summary_table_row + 1
                
         
        
                
        
            End If
            
            
    
                
        Next i
                
                
          ws.Cells(summary_table_row - 1, 12).Value = total_volume
          ws.Range("P1") = "Ticker"
          ws.Range("Q1") = "Value"
         ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & LastRow)) * 100
         ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & LastRow)) * 100
         ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
         ws.Range("Q4").NumberFormat = "0.00"
         increase_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
         ws.Range("P2") = Cells(increase_index + 1, 9)
         decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & LastRow)), ws.Range("K2:K" & LastRow), 0)
         ws.Range("P3") = Cells(decrease_index + 1, 9)
         volume_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & LastRow)), ws.Range("L2:L" & LastRow), 0)
         ws.Range("P4") = Cells(volume_index + 1, 9)
        



Next ws

End Sub



this part was done with the help of ASK BCS Assistance App (later on was edited to add ws. in front of the range to get the resluts of summary tables of all sheets)

         Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & LastRow)) * 100
         Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & LastRow)) * 100
         Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & LastRow))
         Range("Q4").NumberFormat = "0.00"
         increase_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
         Range("P2") = Cells(increase_index + 1, 9)
         decrease_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
         Range("P3") = Cells(decrease_index + 1, 9)
         volume_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
         Range("P4") = Cells(volume_index + 1, 9)

         


