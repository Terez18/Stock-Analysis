# VBA Challenge

## Overview of project: 
### Purpose:
The purpose of this project is to analyze stocks using the most efficient code possible. 
### Bacground:
In the independent learning module code was created to find the total volume and the return on several stocks in the years 2017 and 2018. While the code was able to retrieve all the relevant information, it was not very fast. In order to improve the efficiency of the process of retrieving the information from the stocks worksheets the process of refactoring was used. A significantly faster code was created to retrieve the same information. While both codes achieved the goal and retrieved the relevant information, if a much larger sheet of data was used with many more stocks, then the code that was able to retrieve the information faster would be very helpful. The speed at which the process is happening would become important. 

---

## Results
The time to retrieve the information for 2017 using the old code was 0.734375 seconds and the new code was 0.21875 seconds. The refactored code is significantly faster.

### Image of Time Stamp for 2017 Old Code
![Old_Code_2017.png](Resources/Old_Code_2017.png)

### Image of Time Stamp for 2017 New Code
![VBA_Challenge_2017.png](Resources/VBA_Challenge_2017.png)

---

The time to retrieve the information for 2018 using the old code was 0.7539063 seconds and the new code was 0.2109375 seconds.The refactored code is significantly faster.

### Image of Time Stamp for 2018 Old Code
![Old_Code_2018.png](Resources/Old_Code_2018.png)

### Image of Time Stamp for 2018 New Code
![VBA_Challenge_2018.png](Resources/VBA_Challenge_2018.png)

---

## Results (Cont'd):
- After several repetitions, the code ran even faster taking around 0.16 seconds for the refactored code to retrieve the information. 
- The times taken above are including the formatting in the new code that was faster and without the formatting of cells in the old code that took longer. 
- The VBA file includes several Macros. The last one is the refactored fast code. It is called: Sub AllStocksAnalysisRefactored()
- The old Subroutine that took longer is the one called: Sub yearValueAnalysis()

### Link to VBA and Excel file:
![VBA_Challenge.xlsm](VBA_Challenge.xlsm)

---

### Description of Code
- The reason for the significant reduction in time for the new code is that the computer only went through the worksheet one time and was able to retrieve all the information it needed. The old code was originally created to retrive information about one ticker from one year and then it was expanded to include multiple tickers and two years. For the original goal of one ticker the code was fine. But once more tickers were added, it became inefficient. So instead of looping through the whole worksheet every time a new ticker was analyzed, the refactored code went systematically through the worksheet and took the information from each ticker as it reached it. 
- tickerIndex variable was created and assigned to start at zero.
- Three new output arrays were created so that every time the computer reached the relevant infomration in each ticker, the infomration would be collected for total volume, starting and ending prices. 
- The total volume was set at zero for each new ticker and then increased for the current ticker
- The ticker was evaluated to see if it was the same or the next ticker. This was done in order to determine the starting price and the ending price for the current ticker. Those are used later to calculate returns. 
- Once the ticker changed to the next ticker, the same process repeated for the new ticker until all tickers were done. 
- The output of the tickers, ticker volume, and returns were assigned to the table in the All Stocks Analysis Worksheet
- Formatting of the All Stocks Analysis Worksheet was completed to make the information easier to see and understand. the headers were in bold, a line was added under the headers the numbers were formatted so it is easier to read them and the column size was set to auto fit the correct size. Colors green and red were added to make it easier to see which returns were above or below zero. 
- A message box was created to provide the time it took to complete the code 
- Some of the code for example the table formatting, the timer were used in the old code, and some parts of the code were similar but were updated to allow for this code to work for multiple tickers efficiently. For example, the output at the end was similar but was changed to be tickers(i) instead of ticker. This is so that each time the ticker changed the new output corresponding to the new ticker would be retrieved. 
	
---

## Summary

### Advantages of refactored Code in this VBA Script
This Assignment focused on refactoring code. In this case the purpose was to improve the code to make it run faster. Make it more efficient so if this code was used for a larger number of stocks, it would still work relatively quickly and efficiently. While the differences in the amount of time it took the old code vs the new code to run may seem small and even negligable for this specific assignment, if we took this code and applied it to a large amount of data the time difference would very quickly add up to a large difference. 
The advantages of using the refactored version in this specific code are:
- Time saving: the code takes less time to run. As seen in the images provided. 
- Less work for the computer since it only goes through the sheet once instead of multiple times
- Updated code that is more relevant for the updated purpose. Shifting from retrieving stocks information for one stock for one year to retrieving information for multiple stocks for 2 years. 
- Opportunity to make code more clean and organized. 

### Disadvantages of refactored Code in this VBA Script
The main drawback of refactoring is that no new information was retrieved using the new code so you could say it is a "waste of time" or in other words, if there is limited time available one significant disadvantage of refactoring is that it takes more time!
Disadvatages of using refactored code:
- More time spent for the same information retrieved about the stocks
- In this specific scenario, the running time saved was less than one second so after spending the time to improve the code, the return on investment was a very small change in the time it takes to run the code. This is because the amount of data in this project is not very large. 


### Advantages of Refactoring in General
Refactoring is good practice in general and the advantages are:
- It keeps software more relevant and up-to-date 
- It can make software easier to understand if it is simplified
- Simpler code makes it also easier to find bugs or inefficiencies down the line
- When the software is used for a long period of time, it is helpful to improve the design of the software so you can navigate through it more easily
- It can save time running the code and it unloads the computer, so it can have more power to address other tasks

### Disadvantages of Refactoring in General
Resources are limited so the main disadvantages of refactoring are:
- Refactoring may exceed time alotted for a project
- Refactovring may exceed budget alotted for a project
- If not conducted properly it could introduce bugs or issues into code that was previously working even if not in the most efficient manner. 






