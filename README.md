# stock-analysis

##Purpose

The purpose of this project is to analyze all relavent stocks to our customer for the entire year, and see what their return and total volume look like. Based on this data, we can recommend to the end customer what stocks they should look in to, versus what stocks they should avoid.

The customers originally wanted to invest in DAQO, but upon analysis, it looks like they should try to avoid that stock. The returns are not great for this stock for two years straight. They will want something that can guarantee them a return, even if it isn't a very high return.

##Results

<img width="400" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/95941301/147421304-3f965e02-234e-427c-9acb-fea1d8c0e916.png">

In 2017, we can see that a majority of the stocks are doing pretty well in returns. But, if we carry on into the next year...

![Returns Chart](https://user-images.githubusercontent.com/95941301/147421367-7ac952e4-2cb4-4896-940e-c9ab99d7d822.png)

A ton of them fell! Two stocks have still maintained good returns, and those are ENPH and RUN. We might want to advise the customers to invest in those stocks.

Let's take a look at a part of the code.

```
    For i = 2 To RowCount
    
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'Check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            End If
    
    Next i
```

This is where most of the magic is happening. What we are essentially doing is reading through our main worksheet for a given year, and very quickly taking a look at each ticker. If our ticker is the same one we are looking for, we are going to grab all the relavent information we want from it, which includes the very beginning price, the very end price, and the volume of it. We keep adding up the volume to get a total volume for that ticker!

When we have all of our information stored for each respective ticker, we move onto the next segment.

```
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

For each relavent ticker, we are calculating all the information we need, or we are just filling in the already made value. This all gets outputed to a seperate sheet, where we can easily view the results and better understand everything we were just overwhelmed with!

Now, let's take a look at the run-time for each year.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95941301/147421462-be149451-3489-4ee7-b94e-df018e6e1e0c.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/95941301/147421464-3917c8b5-387f-4b94-81f1-a6c0b0c59813.png)

Each run-time is the exact same, and it happens in the blink of an eye. This performs better than the other code by milliseconds. It might not seem like much, but sometimes those few milliseconds can go very far depending on the use case, especially if you have a very large data set!

##Summary

The advantages to refactoring code are ever present, and can make certain tasks run much faster. This can be very useful if you are on a time crunch and don't want your data analysis to take longer than needed. The only disadvantage I can think of is when a refactored code doesn't make an apparent difference, and you may have spent uneeded time optimizing the code.

Our refactored code runs better by a big margarin! It is faster than the older code and is almost done in the blink of an eye. This plays a big part if the dataset grows to be something enormous, as the older code would take much longer to loop through than the refactored one. Another advantage to the refactored code is that it automatically formats the output when it finishes looping through!

I cannot think of any disadvantages of the refactored code, but one disadvantage of the older code was having to format it seperately after done. We reduced the amount of work needed to be done by the end user. The older code was much slower as well, almost taking a whole second at times to analyze the data.
