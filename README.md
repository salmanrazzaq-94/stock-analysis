# Stock Analysis
Steve is a financial analyst doing stock research for his parents who are interested in investing in Green Energy stocks. 

## Overview of Project
Steve's parents have put all of their money into DAQO New Energy (DQ) and Steve is worried about the lack of diversity in their portfolio. He has provided historical green energy companies stock volumes and prices for 2017 and 2018 in Excel format. Steve has requested help in quickly analyzing the different companies stock performance over the two year periods.

### Purpose
##### Client Perspective
Analyze DAQO stock performances versus the other 11 green energy companies stock performance. Provide analysis and recommendations.

##### Analyst Perspective
Utilize an existing VBA script and refactor the script to be more efficient. 

### Results
#### Analysis
Historical green energy stock prices were summarized for trading volume and return percentage with the following results:

![2017](resources/VBA_Challenge_2017.png)

![2018](resources/VBA_Challenge_2018.png)

DAQO (DQ) performed well in 2017 with an overall return just under 200%. However the volume of stock traded was by far the lowest of the peer group at 35.8M shares. 2018 saw a much higher volume of DAQO stock traded but had a negative 62.6% return. Total return over 2017 and 2018 was 17.8%.

Diversifaction of the stock portfolio is recommended with ENPH and RUN stocks being good candidates with high volumes and positive returns in both 2017 and 2018.

##### Code
Original VBA code had nested for loops that would run through the entire stock data, calculate and write the individual stock's performance and then repeat the process 11 more times:

    'set up stock ticker array
        Dim tickers(11) As String

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

        'initialize variables for starting and ending prices
        Dim startingPrice As Double
        Dim endingPrice As Double

        Worksheets(yearValue).Activate

       'Find number of rows (before both loops)
       RowCount = Cells(Rows.Count, "A").End(xlUp).Row

       'loop through tickers

        For i = 0 To 11

            ticker = tickers(i)

            totalVolume = 0

            'loop through rows

            Worksheets(yearValue).Activate

            For j = 2 To RowCount

                'calc volume
                If Cells(j, 1).Value = ticker Then

                    totalVolume = totalVolume + Cells(j, 8).Value

                End If

                'set Start Price
                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

                    startingPrice = Cells(j, 6).Value

                End If

                'set End Price
                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

                    endingPrice = Cells(j, 6).Value

                End If


            Next j

        'Output results
        Worksheets("All Stock Analysis").Activate

        Cells(4 + i, 1).Value = ticker

        Cells(4 + i, 2).Value = totalVolume

        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1


        Next i
        
###### Run times
Macro run times for original VBA script:

   ![2017](resources/Orig_2017.png)

   ![2018](resources/Orig_2018.png)

The refactored VBA script utilized a Ticker Index to calculate each stock's volume, start price, end price, and return in a single pass through the data:
 
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
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
        
       tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
        
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
        'End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
                
                
            '3d Increase the tickerIndex.
        
        
            tickerIndex = tickerIndex + 1
     
        'End If
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For j = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + j, 1).Value = tickers(j)
    
        Cells(4 + j, 2).Value = tickerVolumes(j)
    
        Cells(4 + j, 3).Value = tickerEndingPrices(j) / tickerStartingPrices(j) - 1
        
    Next j
    
###### Run times
Macro run times for refactored VBA script:

  ![2017](resources/Refactored_2017.png)

  ![2018](resources/Refactored_2018.png)

The refactored VBA script ran dramatically faster than the original script.

### Summary
##### Refactoring VBA code Disadvantages
If the original code works, why spend the time and energy to refactor?

##### Refactoring VBA code Advantages
*  Leverage the existing code to improve it and provide better results
*  Improve efficiency of analysis and macro times
*  Ability to add flexibility to VBA script to better handle future changes in analysis needs

The refactored VBA script not only performs faster but will allow for more flexibility going forward. A larger set of stock data will be able to be analyzed quickly with limited changes needed to the script. Processing times for the original VBA script would have quickly grown with additional data whereas the refactored script will be able to handle much larger data sets with limited increases in processing time.

While getting the refactored VBA to run correctly was challenging at times, the end result is a much quicker and efficient analysis method. Debugging the code enabled better learning that will be remembered.
