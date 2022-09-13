# stock_analysis
# Stock Analysis
##### Automating the analysis of multiple renewable energy stocks using VBA
---

## Overview
This report will analyze the data for 2018 and 2017 of 12 renewable energy stocks including Atlantica Yield, Canadian Solar, Daqo New Energy Group, Enphase Energy, First Solar, Hannon Armstrong Inc., Jinko Solar, SunRun, SolarEdge, SunPower, Terraform Power, and Vivint Solar. The goal is to automate the formatting and analysis using VBA in order to output the yearly return and total volume for each stock. After creating the first draft of working code (Module 1: draftCode), I decided to *refactor* it to come up with a **design pattern** (Module 2: VBA_Challenge) with improved code performance that could be used on any stock data.

## Analysis 
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*Please, refer to the VBA_Challenge.xlsm in the repo specifically to the All Stocks Analysis worksheet. If you want to take a look to the entire code, open VBA and select the module called VBA_Challenge. There is only one Macro called "AllStockAnalysisRefactored".*

I first created an array for all the stocks, so that each one of them could be addressed with an index. 

![Challenge2Exhibit4](https://user-images.githubusercontent.com/83378141/119200473-009cd300-ba5b-11eb-985a-21bee79ee2e2.png)

Moreover, in order to come up with an automated analysis of the stocks I used a nested for loop that went through all the data, starting from the second row to the last one, storing the necessary information into arrays (total volume, first closing price, and last closing price) already initiated at the beginning of the macro. 

![Challenge2Exhibit3](https://user-images.githubusercontent.com/83378141/119200245-8d935c80-ba5a-11eb-9071-b0d85923326f.png)

As you can see in the lines of code above, while the program was looping through all the rows it began storing values into the arrays. Beginning with the total volumes for each stock, I used the command `If Cells(i, 1).Value = ticker Then` to check if the first column had the ticker we needed. If it did, I then did `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value` to store its volume in column 8 and keep adding to it everytime it encountered a row that was of the same stock. 

Furthermore, to check if the row was the first row for that stock I now needed two conditionals `If Cells(i, 1).Value = ticker And Cells(i - 1, 1).Value <> ticker Then`. If these conditions applied, it then went on and set the value in the 6th column (closing price for Jan 1st of that year) to become the first closing price of the stock. Which I used later to calculate the return. 

And lastly, I needed the last closing price of that stock (Dec 31st). In order to do that now I needed two conditionals `If Cells(i, 1).Value = ticker And Cells(i + 1, 1).Value <> ticker Then`. I needed to know that the row was for that stock and that the next row was not from that same stock, checking that first column as well. After storing the value of the last closing price (`tickerEndingPrices(tickerIndex) = Cells(i, 6).Value`), I needed to increase the ticker index so that it went on to the next stock and repeated the entire process for each one of them. 

As the data I wanted to output was already saved into arrays, I only had to create a for loop to do it. Then, it was just a matter of formatting the data so that it showed a green highlight if the return was greater than 0 and red one if smaller than 0 to allow for a better interpretation of the findings.

![Challenge2Exhibit5](https://user-images.githubusercontent.com/83378141/119203477-1f9e6380-ba61-11eb-86d6-6a6fb2a7a83d.png)



## Results

![Challenge2Exhibit2](https://user-images.githubusercontent.com/83378141/119198757-ddbcef80-ba57-11eb-86bc-91491bb47286.png)
![Challenge2Exhibit1](https://user-images.githubusercontent.com/83378141/119198611-9cc4db00-ba57-11eb-88e5-ef5b0d346ace.png)

As it can be seen just by glancing through the findings, the 2017 performance was considerably better as there is a lot more green to be seen in the table. This means there are more returns that came out positive. This is specially true for stocks such as Daqo, that in 2017 came out with a return of 199.4% and in 2018 of -62.6%. Similarly, this same happened for most of the stocks including Atlantica Yield (ticker: AY), Canadian Solar (CSIQ), First Solar (FSLR), Hannon Armstrong Inc. (HASI), Jinko Solar (JKS), SolarEdge (SEDG), SunPower (SPWR), and Vivint Solar (VSLR). 

In conclusion, even though the vast majority of them performed well in 2017, most of them performed poorly in 2018. Definitely, the top picks would have been Enphase Energy (ENPH) or SunrRun (RUN), which were the only green energy stocks that gave high returns for both consecutive years. The best investment would have been to buy a lot of ENPH stock at the beginning of 2017 and hold it through 2018 as it had 129.5% and 81.9% return, respectively. Nevertheless, money could also have been made by buying a put option on TERP at the beginning of 2017 as it returned -7.2% and then in 2018 another -5%. 

### Execution Times and Code Performance

As it was mentioned in the overview of the project, I first created a VBA script with multiple macros that analyzed the stock data and formatted the results. It did the job, however, I then thought I could refactor the code to be able to do everything in one single macro. This would not only decrease the execution time, but it would also serve me as an efficient design pattern that I could share and use with different stock data. 

Before refactoring, my "draftCode" script in module 1 was running the analysis in 0.97 seconds without even including formatting as I had that done in another macro. The code execution time decreased significantly after I refactored. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/83378141/119248253-e8fa4300-bb5d-11eb-8db7-459de88b40a6.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/83378141/119248258-ed266080-bb5d-11eb-8181-c9fff2046f4a.png)

As you can see, the code run time for both years decreased almost by half to 0.57 seconds. Even though this might not be a huge difference in this case, when we work with significantly larger datasets code performance and execution times will be more important to us. 

## Summary

Refactoring code has its pros and cons. Just like proofreading, going back and reevaluating what we wrote is a key step for better results. One pro would be that it makes the script easier to understand. A clean code is more versatile which is a huge advantage as code is always shared. On the other hand, the disadvantages behind refactoring revolve around the fact that it is just time consuming. Some people even argue that it is useless, as you are working on a body of code that already works and gets the job done.  

In this project, this could be easily seen as the code performance and execution times relatively got better. Refactoring the code that did the analysis of the stocks definitely made the script more coherent. Now, I didn't have the iterations in different `for loops` or the formatting in another subroutine. Instead, I could just run the refactored macro and everything would be done at once. Even though the advantages were many, it took me a lot of time to come up with the refactored version of the code. I needed one single `nested for loop` to do the job and store the data I needed into arrays rather than simple variables. But even though it was time consuming, I believe it was worthwhile as now I have a body of clean code that I can easily edit in order to analyze different stock data. The goal should be to always strive to create a design pattern that is versatile, clean, easy to understand and shareable. 

In conclusion, the original script had the advantage that if the code was wrong at any point, it would only affect that macro and function. The others would still work. Meanwhile, the disadvantages were that it was a very long and all-over-the-place script. On the other hand, the advantage of the refactored script was that it was more efficient and organized, making the execution times a lot shorter and the script easier to follow. 
