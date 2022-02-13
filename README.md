# Stock Analysis with VBA
## Overview of Project
### Purpose
Steve was impressed with the first workbook that was created that allowed him to analyze an entire dataset in one push of a button. As a result, he wishes to do more research and expand the dataset to include a larger dataset over the years. I have been tasked with editing, or refactoring the VBA code to loop through all the data one time to collect the same information at a much more proficient speed. I will then determine if refactoring the code successfully make the VBA script run faster. When refactoring the code, I will be making the code run more efficiently, with fewer steps, less memory, and making it easier to read.
## Results
### 2017 and 2018 Stock Analysis: Comparison
When looking at the stock results from 2017 compared to 2018, we can see that 2017 had a much better return overall than 2018. In 2017, 11 out of the 12 stocks had a positive return, with DQ having a return as high as 199.4%. Meanwhile, when we take a look at the returns in 2018, only 2 out of the 12 had a positive return where the ENPH and RUN stock had a 81.9% return for ENPH and a 84.0% return for RUN. So overall, if you were to want to invest in a stock in 2018, you would have been better off buying in on RUN. In 2017, you would have still been getting a better return on ENPH with a 129.5% return. In conclusion, 2017 was a better year on returns than what 2018 produced.

![image](https://user-images.githubusercontent.com/97328622/153737959-95c545d2-c5b3-461d-8613-92e549021160.png)
![image](https://user-images.githubusercontent.com/97328622/153737969-80f3a766-48fe-43ac-981d-0d33be8d9eb9.png)

First Photo: (VBA_2017_data_results.png), Second Photo: (VBA_2018_data_results.png)

## Original Run Times vs. Refactored Run Times
Original run time of 0.765625 seconds for 2017 and 0.7734375 seconds for 2018.

![image](https://user-images.githubusercontent.com/97328622/153737718-f2a11b4b-d99d-4af9-b7f8-5dc588170e3f.png)
![image](https://user-images.githubusercontent.com/97328622/153737753-d426f30e-f230-42c7-8845-fef66aa9a38f.png)

First Photo: (Original_2017_run_time.png), Second Photo: (Original_2018_run_time.png)

Refactored run time of 0.0859375 seconds for 2017 and 0.0859375 seconds for 2018.

![image](https://user-images.githubusercontent.com/97328622/153738159-af4294aa-d194-4346-9382-8d430c37c006.png)
![image](https://user-images.githubusercontent.com/97328622/153738188-bc92ecec-78e1-4397-9d1b-7a3eeb166f86.png)

First Photo: (VBA_Challenge_2017.png), Second Photo: (VBA_Challenge_2018.png)

The refactored code decreased the run time of the VBA code by 0.6796875 for 2017 and 0.6875 for 2018. This shows that the refactored code worked more efficiently than the original VBA code which means the refactoring of the code was successful.
## Summary
### Advantages of refactoring code
The advantage of refactoring code is how efficient it can be for, not only the data, but also for successfully creating a code that takes less steps, uses less memory, and improves the logic of the code to make it much easier for future analysts and users to read. Refactoring can also help to debug an original VBA script that is having trouble running a clear and concise dataset. Finally, refactoring a code allows a fresh pair of analytical eyes to put their own spin and insight on the code which can potentially help to improve the code even further.
### Disadvantages of refactoring code
The disadvantage of refactoring code is that refactoring an existing code that already needs to be debugged may further bug that code even further if there isn't a careful set of analytical eyes editing the existing code. This may lead to even further bugs and a more complex VBA script that may need to be refactored even more. This would lead the refactoring to take more time than it needed to be. Each Analyst and programmer is different and brings in their own approach, which means that some individuals may not be able to collaborate efficiently.
### How the possible advantages and disadvantages effected my own process
I was able to experience the advantages of refactoring an existing code. Due to the results listed above, I was able to successfully refactor the code to become more efficient and run quicker than the original VBA script. The issue I faced in refactoring the existing code from the previous programmer is that its difficult to determine the original direction that the individal was trying to go. I had to take time to study each section, piece by piece, and ultimately change most of the code, which ended up being time consuming. Overall, the advantages of refactoring an existing code was heavily in my favor and, as a result, left the refactored code being more efficient than the original.
