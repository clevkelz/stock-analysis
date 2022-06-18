# Using Visual Basic (VBA) to Evaluate Green Companies Stock Returns and Volume for 2017 and 2018
## Overview of the Project
The return on stoc from 2017 and 2018 for the following 12 green energy (also referred to as "Green Companies") was reviewed to show the differences in daily volume and return among these companies.  This analysis was also performed to show how VBA can be used to analyze and make conclusions from a large dataset, in this case, daily stock prices for 12 Green Companies over the span of two years.

|Company Name|Stock Ticker|Website|
|----|-----|-----|
|Atlantica Sustainable Infrastructure PLC|AY|https://www.atlantica.com/web/en/|
|Canadian Solar Inc.|CSIQ|https://www.canadiansolar.com/|
|Daqo New Energy Corp|DQ|http://www.dqsolar.com/|
|Enphase Energy Inc|ENPH|https://enphase.com/|
|First Solar, Inc.|FSLR|https://www.firstsolar.com/|
|Hannon Armstrong Sustainable Infrastructure Capital, Inc.|HASI|https://www.hannonarmstrong.com/|
|JinkoSolar Holding Co., Ltd|JKS|https://www.jinkosolar.com/en/|
|Sunrun Inc|RUN|https://www.sunrun.com/|
|Solaredge Technologies Inc|SEDG|https://www.solaredge.com/us/|
|SunPower Corporation|SPWR|https://us.sunpower.com/|
|TerraForm Power|TERP|https://www.terraform.com/|
|Vivint Solar Inc|VSLR|No longer in existence - https://www.nytimes.com/2020/07/06/business/energy-environment/sunrun-vivint-solar.html|

## Results
### Summary of Stock Returns and Daily Volume
Stock returns for Green Companies were strong in 2017, with only one of the 12 companies experiencing negative returns; however, this pattern reversed sharply in 2018, when only two of the companies had positive returns.  Total daily volume for these companies was generally higher in 2017 than in 2018 with the notable exceptions of Daqo New Energy Inc., Enphase Energy Inc., and Sunrun Inc. 

![image](https://user-images.githubusercontent.com/106293233/174418906-99ce9090-668c-4ef5-9861-a5749fb54a1d.png)

![image](https://user-images.githubusercontent.com/106293233/174418951-9e3b4b9f-5d62-4d2f-b93a-fd9c53c6c172.png)

The stock market experienced large losses in the last quarter of 2018, which is likely the largest contributor to the fall in Green Company Stock Prices in 2018.  https://www.pbs.org/newshour/economy/making-sense/6-factors-that-fueled-the-stock-market-dive-in-2018

### Differences Between Origingal and Refactored Scripts

### Running Time
The original analysis was completed in the file called "green_stocks.xlsm" - https://github.com/clevkelz/stock-analysis/blob/main/green_stocks.xlsm - while the refactored code is inthe file "VBA_Challenge.xlsm" - https://github.com/clevkelz/stock-analysis/blob/main/VBA_Challenge.xlsm.  Under similar operating conditions.   the running time for the updated scipt was slighly faster than that for the original script.  

_Running Time for Original Script_
![image](https://user-images.githubusercontent.com/106293233/174420050-8828169f-96a6-46e8-aa6d-d4eb352029d5.png)

_Running Time for Updated Script_
![image](https://user-images.githubusercontent.com/106293233/174420216-a1474c3e-ee72-4cef-8585-8992c43deff6.png)

### Formatting
The two scripts returned the same values for the analysis; however, the formatting on the results was in a different subroutine in the original spreadsheet.  An additional macro had to be run to commas in the values in the Total Daily Volume column, to format the Returns column as a percent rather a long decimal, and to highlight positive returns in green and negative returns in red.  The updated script included the formatting in the same macro which negated the need to run two separate macros to get the same results in the refactored code.  

![image](https://user-images.githubusercontent.com/106293233/174420923-cc332a90-2190-4440-a9ec-8c9df7600be7.png)

![image](https://user-images.githubusercontent.com/106293233/174421094-3f8c5d07-62fb-425b-8dac-ed803885010b.png)









