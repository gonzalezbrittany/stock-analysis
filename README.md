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

