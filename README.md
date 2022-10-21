# Stock-Analysis
Module 2 Challenge involving the analysis of various green stocks.

# Green Energy Stock Analysis

## Project Overview
In this Module 2 Challange, we are assisting our friend Steve in building a dynamic excel workbook that quickly analyzes returns on a portfolio of green stocks. 

### Purpose
The purpose of the project, more specifically, was to analyze data related to green energy stocks in 2017 and 2018 and see how daily trading volume and annual returns were impacted. 

Steve is looking to do a more robust analysis for his parents. They are interested in investing in DAQO stocks. DAQO (ticker sign "DQ") is a company that makes silicon for solar panels. In order to determine the potential quality of the investment, Steve wants to run an analysis on DQ's stocks performance over the period between 2017 and 2018 in comparison to other green-energy stocks. The ultimate goal of the anaylsis will be to determine whether or not DQ stocks specfically are worth investing in relative to other green stocks in the index.

### Background
In this project we are using VBA to code and automate the analysis process for Steve. In the analysis we are using various VBA tools sucj as ***for loops***, ***conditionals***, and ***conditional formatting***. 

In our initial attempt to do data analysis for Steve, we ran VBA code that pulled data in a less than efficient manner (more to come in results section). The challenge required us to ***refactor*** the initial code to have it run more efficiently. The ultimate goal of the code refacotring is to ensure that the integrity of the code is strong, it runs as efficiently as possible in its current state, and that it contains enough descriptive text to be legible and refactorable for other analysts to contribute. 

## Results
### Stock Returns Analysis
When running the analysis. We saw that there were drastic differences in the performance of green stocks in 2017 and 2018. In 2017, green stocks saw strong annuakl returns with 10 of the 11 selected stocks generating positve returns and "DQ" specifically generating the highest overall returns (199%) of the entire basket of securities. Additionally, in 2017 there were 4 of 11 stocks (~36%) that generated over 100% returns during the year. 

2018, however, was a different story. in 2018, only 2 of the 11 stocks in the analysis were able to generate positive returns for investors. Of the 9 companies that generated negative returns, "DQ" generated the highest losses with a -64% return. 

<p align="center">
<kbd>   
<img width="50%" alt="2017 Return Analysis" src="https://user-images.githubusercontent.com/115036844/197091579-38bb41f5-03de-4590-989f-dd0428f09793.png"> <img width="50%" alt="2018 Return Analysis" src="https://user-images.githubusercontent.com/115036844/197091595-3edf3fa7-fda1-4eec-ab2f-d47ec4ca20ea.png">
</kbd>
</p>

There are various reasons why there could be such drastic difference between the 2017 and 2018 data. Steve should first look at the economic and overall market conditions in the two years including interest rates, federal funding, and earnings. Furthermore, Steve should also look into any regulatory changes that might have impacted the performance of securities within the basket. The EPA and SEC both can have significant impact on various finanicing mechanisms for companies across renewables and therefore need to be researched. 

It is important to note that "DQ" specifically seems to be the most volatile of all the stocks in the basket. In 2017, the stock was the best performing and in 2018 the stock was the worwst performing. This could be due to the overall interest rate environment but more likely has to do with structral risks in the business. If Steve's parents are looking for high potential returns and have an appetite for risk, 
  
### Code Refactoring Analysis
Both runs of the code generated the same output (as demonstrated above). The refactored code runs calculations from the original dataset provided and generates return data and trading volume data on a new worksheet named All_Stock_Analysis. The reason for running two seperate sets of code is primarily to demonstrate how refactoring can ultimately lead to better process outcomes and efficiency.

The largest differnce between code bases is the use of nested loops versus multi-dimseinal arrays. 

In the intial code run, we utilzied nested loop code which caused us to switch between workbooks as we looked to generate return data. When working simply on a smaller data subset (the specific tickers that we were looking at) the code functioned effectively but inneficiently. With the introduction of arrays we were able to store the data as elements within the array and therefore improve execution speed and efficiency. 

The code below shows the refactored code that was used to improve efficiecny specifically in the context of calling tickers. 

    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

The ability to access arrays with a single variable tickerIndex was the key driver in increasing efficiency. In this case, code stored all elements in arrays before switching to another worksheet as opposed to iterative switching in the origibal code.

The below shows the ***original*** codebase and how the iterative loop cycle would call between workbooks to generate data outputs. 

    '4) Loop through the tickers.
    
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0

    '5) Loop through the rows in the data.
        'Activate Data Worksheet
        Worksheets(yearValue).Activate
        
        For j = rowStart To RowCount
        
    
    '5a) Find the total volume for the current ticker.
    
            'Identify ticker
            If Cells(j, 1).Value = ticker Then
                
                'increase ticker totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
    '5b) Find the starting price for the current ticker.
    
            'Identify first row of ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
                
            End If
            
    '5c) find the ending price for the current ticker.
    
            'Identify last row of ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
    
    
    '6) Output the data for the current ticker.
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker header
        Cells(i + 4, 1).Value = ticker
    
        'Sum of Volume
        Cells(i + 4, 2).Value = totalVolume
    
        'Return Value
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
        
        
 ### Improved Efficiency
 Ultimately, by making the changes described above, we were able to meaningfully impve the speed at which the code could be run. 
 
 Code before refactoring (Module 1). |  Code after refactoring (Module 2).
:------------------------------------------:| :-------------------------------------:
2017 Before Refactoring (click to enlarge).  | 2018 Before Refactoring (click to enlarge).	
<img width="262" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/115036844/197101250-d04cd2a9-7fe1-4d06-8eab-2d1daf7bd936.png"> | <img width="262" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/115036844/197100546-ecad57e1-468c-4ebd-8b5f-1f476997f761.png">
The code in a nested loop is switching back and forth between worksheets. | Code stays in the same loop, gathers all data and stores it in arrays. 
 2017 After Refactoring: |  2018 After Refactoring:
<img width="259" alt="VBA_Challenge_2017_Refactored" src="https://user-images.githubusercontent.com/115036844/197100990-6a66579c-6365-43a0-af37-cfa44b9e0f39.png"> | <img width="266" alt="VBA_Challenge_2018_Refactored" src="https://user-images.githubusercontent.com/115036844/197101037-121311fe-f0ee-435a-80f1-7a63ac2ccbd3.png">

After our modification to the code, the analysis was able to run approximately 5 times faster than it had previously. 

## Summary
## Advantages of refactoring code

The obvious advantage of refactoring code is the imporoved efficiency and speed. This specifically means that the code can handle larger data sets without potentially running into issues. An 5x increse in speed to execution can be importsnt when there is a possibility to increase the overall size of a dataset. Additionally, the improved efficency will likely make it easier for future contributors to make modifications to the existing code.  

### Disadvantages of refactoring code

Refactoring always poses the risk of breaking existing code that is functioning correcrtly (although maybe not at its most efficient). Refactoring can potentially lead to errors or omisisons that might otherwise not have been made.  Whenever refacotiring occurs, it is a good idea to save initial versiions of code to preserve the integrity of the initial code run.  

## Additional Code

### Refactored Code (in full)

```
Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Title Analysis
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
    
    'Count the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index to reference proper ticker in the arrays.
    Dim tickerIndex As Integer
    'Initiate tickerIndex at zero.
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    
    '2a) Create for loop to analyze each ticker in the array.
    For tickerIndex = 0 To 11
    'Initiate each ticker's volume at zero.
    tickerVolumes(tickerIndex) = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the first row for current ticker, set starting price.
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End If
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                'if it is the last row for current ticker, set ending price.
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End if
            End If
            
        '3d) Check if the current row is the last row with the current ticker.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                'if it is, increase tickerIndex to move on to next ticker in array.
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of Volume
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'ReturnValue
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    
    Next i
```


### Original Code (in full)

```
Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
'1) Format the output sheet on All Stocks Analysis Worksheet
    'Activate "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate
    'Title Analysis
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a Header Row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2)Initialize an array of all tickers.
    
    'Declare an array with 12 string elements
    Dim tickers(12) As String
    
        'Assign tickers to an element in the array
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
        
'3) Prepare for the analysis of all tickers.
    '3a) Initialize variables for the starting price and ending price.
    
        'Creating a Variable for Starting & Ending Price
        Dim startingPrice As Double
        Dim endingPrice As Double
    
    '3b) Activate the data worksheet.
        
        Worksheets(yearValue).Activate
        
    '3c) Find the number of rows to loop over.
        
        rowStart = 2
        'DELETE: rowEnd = 3013
        'rowCount code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
'4) Loop through the tickers.
    
    For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
'5) Loop through the rows in the data.
        'Activate Data Worksheet
        Worksheets(yearValue).Activate
        
        For j = rowStart To RowCount
        
    
    '5a) Find the total volume for the current ticker.
    
            'Identify ticker
            If Cells(j, 1).Value = ticker Then
                
                'increase ticker totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
    '5b) Find the starting price for the current ticker.
    
            'Identify first row of ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
                
            End If
            
    '5c) find the ending price for the current ticker.
    
            'Identify last row of ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
    
    
'6) Output the data for the current ticker.
        'Activate Output Worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker header
        Cells(i + 4, 1).Value = ticker
    
        'Sum of Volume
        Cells(i + 4, 2).Value = totalVolume
    
        'Return Value
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
   
End Sub
```


