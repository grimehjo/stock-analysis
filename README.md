# stock-analysis

In this exercise I analysed different green stocks using VBA.

## Deliverable 2: Written Analysis of Results

### Overview of Project: Explain the purpose of this analysis.

In this exercise I analysed different green stocks using VBA for an individual named Steve, whose parents are interested in investing money into environmentally friendly stocks.

I first analysed the stocks of only one company with the ticker symbol of DQ. DQ was the stock the family were first interested in investing, but unfortunately the returns were discovered to be pretty bad for this stock after doing an analysis using VBA. Because of this, they decided that they would actually like to look at all the stocks available in the data to see if there are any better options.

Because of this, I then created a VBA code to analyse all the stocks at once. I then made the data interactive by letting the user to be able to analyse the stocks by year.


### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

The stock performance was much better in 2017 compared to 2018. In 2017, 11 out of the 12 stocks analysed returned a profit. In 2018, only 2 out of the 12 stocks analysed returned a profit.

2017 Stock Analysis:

<img width="148" alt="Analysis2017" src="https://user-images.githubusercontent.com/80979705/119277191-30c7ab80-bbec-11eb-8fc7-d51863273dca.PNG">

2018 Stock Analysis:

<img width="151" alt="Analysis2018" src="https://user-images.githubusercontent.com/80979705/119277215-4b018980-bbec-11eb-8d1b-9de11a064c77.PNG">

By looking at the returns, it is obvious that 2017 was a much better year for most of these stocks than 2018. 

Original 2017 Stock Analysis Execution Time:

<img width="151" alt="2017OriginalTime" src="https://user-images.githubusercontent.com/80979705/119277282-bba8a600-bbec-11eb-88ce-cdb26a571156.PNG">

Original 2018 Stock Analysis Execution Time:

<img width="158" alt="2018OriginalTime" src="https://user-images.githubusercontent.com/80979705/119277289-c6633b00-bbec-11eb-8683-2ea666b7fb19.PNG">

The code I wrote worked well but it is not very efficient. The code took .844 seconds to process the 2017 stock data, and the code took .836 seconds to process the 2018 data. While this might sound fast, it is actually very slow- especially if we intend to use this code on exponentially more stock data like many real wall street analysts do. 

As this code might want to be used on even more stocks and data in the future, it is important to speed up the execution time to make the script perform the stock analysis faster. One easy way to speed up the execution time of the code, is to refactor the script to limit the amount of nested loops present- which is one of the more time consuming tasks for the program because each loop requires the program to read the entire data sheet from top to bottom again. 

There is a way to avoid this though. We can refactor or restructure the script to perform more quickly and efficiently by making sure the code only needs to loop through the data one time to get all the information it needs. You can do this by breaking down and reorganising your current code. This is what I did with the following results:

Refactored Code 2017 Execution Time:

<img width="156" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/80979705/119277302-dbd86500-bbec-11eb-95fa-60cf6dcbd9e7.PNG">

Refactored Code 2018 Execution Time:

<img width="157" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/80979705/119277311-e72b9080-bbec-11eb-8b7a-4ea628efb1f4.PNG">

As you can see from the above screenshots, the execution time was much faster because the performance was made to be much more efficient.

For example, here is part of the original script before the refactoring. As you can see I have a For loop inside another For loop:

For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       
       '5) loop through rows in the data
       
    Worksheets(yearValue).Activate

       
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           
           '5b) get starting price for current ticker
           
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       
       '6) Output data for current ticker
       
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

### Summary: In a summary statement, address the following questions.
#### What are the advantages or disadvantages of refactoring code?
#### How do these pros and cons apply to refactoring the original VBA script?
