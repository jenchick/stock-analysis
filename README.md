


# Stock Analysis with Visual Basic Applications (VBA)

## Overview of Project

### Purpose
The purpose of this project is to expand a stock market analysis project for a single stock (DQ) to a larger subset of stocks (n=12).  I will also analyze the results of this stock market data from 2017 and 2018, along with refactoring code to determine if the new code positively affects performance of the scripts run to calculate stock performance.

## Results: 
### Stock Performance
As evidenced by the 2017 to 2018 Select Stock Performance graph below, stock performance for 2017 produced better returns than the dismal performance of 2018. Only two stocks, ENPH and RUN produced positive returns in both years.  SEDG had the second best returns in 2017, with minimal losses is 2018.  Although DQ was the highest perfoming stock in 2017, it also had the highest overall loss in 2018 showing great volatility.

![image](https://user-images.githubusercontent.com/102322707/166079568-b36c57c3-ad69-4a4a-bec8-812a360d6bc9.png)

### Script Performance
#### Original Script
The results of the original script are copied below.  For both 2017 and 2018, the code ran in just over 1.03 seconds.  Both runs produced the expected output and formatting.

![AllStocks_run_2017](https://user-images.githubusercontent.com/102322707/167197444-f1370c71-8476-492b-bc6f-6fc6a00a50e2.PNG) 

![AllStocks_run_2018](https://user-images.githubusercontent.com/102322707/167197559-4013a070-05a1-4019-9f78-0f76626f2d0c.PNG)

#### Refactored Script
The results of the refactored script are copied below.  For 2017 and 2018, the code ran in 0.160 and 0.133 seconds, respectively.  Both runs produced the same output and formatting as the original code, but ran much faster by 0.870 seconds for 2017 and 0.897 seconds for 2018.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/102322707/167198080-fca92884-5ff1-4755-950d-525ba376d15e.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/102322707/167198094-138f5329-9834-4070-8db4-7c3dc5eb3af9.PNG)

During the refactoring process, I referred to the "hint" given in the Module in order to increase the volume of the current tickervolumes by using the tickerIndex variable as the index by using the suggested code of:  *tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value*


In addition to refactoring the code, I also corrected the timer output statement from my original script to include proper spacing and sentence structure by use of the following code: *MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year" & (yearValue)*

## Summary: 
### Advantages or Disadvantages of Refactoring Code
In reviewing internet searches on the advantages and disadvantes of refactoring code, contributors to this question on Stackoverflow.com (https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software) summed it up best.  Possible advantages include better quality code, improved code maintenace, clear and precise organized code, and the discovery of bugs in the code. Refactored code could also be more efficient, using fewer step and less memory, along with being more organized and logical.

Possible disadvantages include the time and resources (potentially cost) taken to retest functionality utilizing new code; the bigger the application, the bigger the risk.

### Pros and Cons to Refactoring the Original VBA Script
One of the most obvious positive results to refactoring the original VBA Script is the speed in which the output was generated. The larger the data set, the more time would be saved with the refactored code.  Also, there would be a time savings depending on how many times the script would need to be run.  

Obstacles to refactoring the original VBA script would be the time and energy utilized for the person creating the new code, along with the knowledged needed to create the improved script.   
