# Stock_analysis

## Overview of Project

This analysis was done with the purpose of retrieving and displaying summarized data on the prices and trade volumes of selected stocks.
Our goal was to write code that would make it easier for our user to retrieve specific information about specified stocks, also to present
this data in a way that helps the user make informed decicsions.  The important aspect of the code written is that it returns the expected output in a timely manner.


## Results

The result of the analysis shows in 2017 all but one of the selected stocks had high returns. Compare this to the returns in 2018 where all but 2 stocks had negative returns. 

https://user-images.githubusercontent.com/99847046/158042235-4b5f6e0d-bbd9-4edf-beb7-b4ead2ab56ba.png

https://user-images.githubusercontent.com/99847046/158042318-8208ae0c-4aee-4ee7-933d-eb5ed038d8e6.png

The VBA code was written in a way to make it easy for the user to enter the specified year he was interested in and formatted that a quick glance could readily tell the story
of stock performance in that year.  The original code was refactored to allow the code to be ran more quickly and use less resources to improve delivery effiency. 
Below is the original execution time of the code.  The run time was virtually the same regardless of year chosen.

https://user-images.githubusercontent.com/99847046/158042376-56b7c51c-ef24-474b-aa3c-b2e8921f82b2.png

Compared to the the refactored version of the code that ran at a fraction of a second due to utilizing arrays to limit the number of times the code was looped.

https://user-images.githubusercontent.com/99847046/158042351-c29b8bd6-29aa-468d-8d1c-c8e4b2808dc6.png

## summary

The stand-out advantage of using refactored is the improved speed results are returned.  As small as a time difference of 2/3rds of a second may seem, it made the process feel much smoother and more fluid.  Another advantage is that we already had lines of code that we knew worked and just had to make minor changes to improve the efficiency of delivery where we saw fit. This advantage can work out to be the disadvantage of using refactored code as attempts to adjust code previously written can be tedious and error prone.  
Errors will likely be made especially if the person refactoring the code did not write the original and may even misinterpret the notes left in the code by the original writer.
For example in refactoring the VBA code for this analysis there was some confusion over what it meant to "intialize an array to 0".  More experienced coders might have no issue understanding what is meant but inexperienced coders might overthink and overcomplicate the code. 
