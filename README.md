# VBA Stock Analysis

### Data Overview
This analysis analyzes green energy data for 12 different stocks. The data represents each stock as a “Ticker” variable. Daily stock volumes data was collected throughout the year and recorded in a spreadsheet. This analysis will use macros to analyze provided data. The created macros will display the total daily volume for each ticker for the user requested year. It will also show the percent of return for each ticker by using the daily close price at the start of the requested year and the daily close price at the end of the requested year. This will allow us to see which ticker produced the most return for the year and which produced the least.
After creating a macro that shows which ticker produced the most return and which the least, a follow up analysis was done to see how long it takes to run the created macro. A timer is used in the VBA macro to record run time throughout the macro. To see if the run time can be reduced, the original code was refracted to attempt to reduce the code run time.  

### Results
#### Stock Data Analysis
To find which stock had the most return for the request year, we first asked the user what year of stock data needed to be analyzed used the code below. Once the year was indicated, a data table was created in the All stocks Analysis spreadsheet.
	  
![Picture1](https://user-images.githubusercontent.com/26393180/148663643-16c74ac0-add1-4ce1-b834-9f1d241f30fb.png)
![Picture2](https://user-images.githubusercontent.com/26393180/148663662-ddcb77bc-06be-4e97-8cf8-4bf6058f03e6.png)

To calculate the Total Daily Volumes and Returns for each stock, the below code was used to perform the calculation for the requested year.

![Picture5](https://user-images.githubusercontent.com/26393180/148663707-d9df78ba-d266-4e8b-8447-1da69969ee75.png)
 
The results of the analysis were added to the created data table and are presented below. The return column was then color coded based on if the return from the start of the year to the end of the year was higher (green) or lower (red).


![Picture3](https://user-images.githubusercontent.com/26393180/148663689-3a4fb0b0-676c-46b7-b256-48d2c8e0e453.png)
![Picture4](https://user-images.githubusercontent.com/26393180/148663691-feeb09d6-f852-4515-a029-f2c87d99a668.png)

Our analysis shows that in the year 2017, the ticker DQ had the most return at the end of the year while the ticker TERP had the least return at the end of the year. Our 2018 data table presents a different story, the ticker that had the most return at the end of the year was RUN while the ticker with the least was DQ.

#### Run Time analysis
After confirming our macro works correctly and analyzing the results, the code was refactored to see if we can reduce the run time for the code. The code was edited to allow the data to be looped through one time instead of multiple times. To allow this, we created a “tickerIndex” variable and created output arrays prior to the data rows being analyzed. After running the macro, the analyzed data table results were compared to the original data table results to ensure the edited formulas ran correctly. The output below shows the run time results for the 2018 data for both the original code and refactored code.


![Screenshot 2022-01-08 174537](https://user-images.githubusercontent.com/26393180/148663730-18a26d28-adc7-444a-9d77-a26accf700e1.png)

The run times displayed above show that by refactoring the data so the code loops through one time instead of multiple, the run time was reduced. 

### Summary
#### Advantages of refactoring code
Refactoring a code can be beneficial when it comes to data analysis. The first version of a code may be messy and hard to follow. By refactoring the code, it can be cleaned and reformatted so other users can understand different parts of the code. It may also lead to eliminating repetitive lives and reduce run time. For the VBA script refactored in this analysis, by eliminating repetitive loops and creating an output array prior to the row analysis the run time was able to be reduced. If more tickers are added to the sheet and are also needing analysis, this refactored code will save time when analyzing more tickers.
#### Disadvantaged of refactoring code
Even though there are benefits to refactoring code, there are also disadvantages as well. When looking for a way to refactor code, there is always the possibility that the original code is the cleanest and simplest version possible. It may be a waste of time to try and simplify it further. There is also always the chance that it may not be beneficial to take the time to try and refactor the data. For this analysis we were analyzing a small dataset and were able to refactor the code. But, in doing so the run time was only able to be reduced by a few tenth of a second. Taking time to factor code for a large dataset could substantially reduce the run time but for this smaller dataset it is not very beneficial.  



    Sub AllStocksAnalysisRefactored()
    'Timer start and end time variables are created
    Dim startTime As Single
    Dim endTime  As Single
    
    'User input box is created asking what years data the user would like this module to analysis
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Start time is recorded
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index

        tickerIndex = 0

    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    'a for loop is created to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
  
        Next i
            
        
    'A for loop is used to look over all the rows in the spreadsheet.
    'To ensure we are analyzing the data indicated by the user, the user entered year spreadsheet is reactivated
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        'Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        'If then is used to check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'If then is used to check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
    
        End If

            'Once last row of current ticker is initiated, the ticker is increased by one
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = (tickerIndex + 1)
        End If
    Next i
    
    'Created arrays are looped through to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'End time is recorded, text message box shows how long it took to run module
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

