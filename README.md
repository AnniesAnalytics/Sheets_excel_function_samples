# Sheets_excel_function_samples
The partner file to my SQL file!

# AVERAGE / AVERAGEIF
=AVERAGE(A2:A8)<br/>
=AVERAGEIF(A2:A10, “NY”, B2:B10)

# Arithmetic 
Your spreadsheets can do basic math for you. Just put an = , then the two cells you want to use, separated by your operator. + - * or /        So, =A2-B3    or =A2*B5   Also see, modulo.

# COUNT- A/IF
COUNTA counts the total number of values within a specified range.  COUNT only counts the numerical values within a specified range.

## COUNTIF will count the values within a range as long as they meet a certain condition. 
=COUNTIF(range, criterion)<br/>
=COUNTIF(A1:A9, “food”)<br/>
Instead of inputting the value/string into the “ “ in this function, you can remove the “ “ and instead just type/click a cell which has the string you want to look for.

## =COUNTIFS(criteria_range1, criterion1, [criteria_range2, criterion2, ...])
=COUNTIF(A1:A9, “coffee”, C1:C9, “12/12/2020”)




# CLEAN BLANKS
Excel: 
CTRL+G, click “specials”, “blanks”, all blanks are selected, CTRL + -, “shift cells up”


# Change format of number strings 
Format > number > can change to date, time, money, etc. This allows you to read them correctly in functions

# Conditional Formatting
To change the color and appearance of cells based on their color, <br />
Format>conditional formatting> add rules > it will change their colors

# Convert numbers (currency, temp)
In a new column type =CONVERT(B2,”F”, “C”) =CONVERT(cell you are converting from, “unit from”, “unit to”) then you can apply to the entire column… Now make a new column, copy the values, and paste special (values only) and delete the original column.

# Converting Data

## String to date How to convert text to date in Excel: Transforming a series of numbers into dates is a common scenario you will encounter. This resource will help you learn how to use Excel functions to convert text and numbers to dates, and how to turn text strings into dates without a formula. 
Google Sheets: Change date format: If you are working with Google Sheets, this resource will demonstrate how to convert your text strings to dates and how to apply the different date formats available in Google Sheets. 
## String to numbers
How to convert text to number in Excel: Even though you will have values in your spreadsheet that resemble numbers, they may not actually be numbers. This conversion is important because it will allow your numbers to add up and be used in formulas without errors in Excel. <br />
How to convert text to numbers in Google Sheets: This resource is useful if you are working in Google Sheets; it will demonstrate how to convert text strings to numbers in Google Sheets. It also includes multiple formulas you can apply to your own sheets, so you can find the method that works best for you. 
## Combining columns
Convert text from two or more cells: Sometimes you may need to merge text from two or more cells. This Microsoft Support page guides you through two distinct ways you can accomplish this task without losing or altering your data. It also includes a step-by-step video tutorial to help guide you through the process.<br />
How to split or combine cells in Google Sheets: This guide will demonstrate how to to split or combine cells using Google Sheets specifically. If you are using Google Sheets, this is a useful resource to reference if you need to combine cells. It includes an example using real data. 
## Number to percentage
Format numbers as percentages: Formatting numbers as percentages is a useful skill to have on any project. This Microsoft Support page will provide several techniques and tips for how to display your numbers as percentages. <br />
TO_PERCENT: This Google Sheets support page demonstrates how to use the TO_PERCENT formula to convert numbers to percentages. It also includes links to other formulas that can help you convert strings. <br />
Pro tip: Keep in mind that you may have lots of columns of data that require different formats. Consistency is key, and best practice is to make sure an entire column has the same format. <br />
Additional resources<br />
If you find yourself needing to convert other types of data, you can find resources on Microsoft Support for Excel or Google Docs Editor Help for Google Sheets. 
# Combine
To combine the strings from two cells together, use =CONCAT(A2:B2) <br />
To combine the two strings but add something in the middle, use =CONCATENATE(A2: “ “, B2). What is in the “” will be in between the two strings, in this case a space. Another example of this: <br />
=CONCATENATE(C2, " ",D2, ", ",E2)
# Data Validation
Allows users to add data to a column in a specified format, like in a multiple choice quiz. To use go to the data tab > data validation > choose the cell range > list from a range > type the options (ex. Not yet started, finished). You can also use this to add checkboxes


# FIND 
Will locate specific characters in a string, it is case sensitive. If two pieces of data are separated by a space, then you can use this to find out at which character. =FIND(“ “, A2) Where “ “ refers to the space and A2 is the cell.


# Import a table from a webpage
Importhtml

# LEN
Find the length of a string. =LEN(A2)  where A2 is the cell you want. 

# MOD (Modulo)
An operator (%) that returns the remainder when one number is divided by another.

# MIN/MAX
To find the minimum or maximum value in a dataset, use =MIN(A2:A10) or =MAX(cell1:lastcell)<br/>
MAXIF : The first argument, max_range, is the array over which you are finding the maximum. The second argument (range1) is the array you are checking. The third argument (criteria1) is the value that you are checking for. The inputs in the square brackets are for optional additional constraints.

## =MAXIFS(max_range, range1, criteria1, [range2], [criteria2], ...)
=MAXIFS(D2:D21, B2:B21, "NY")        =MAXIFS(D2:D21, B2:B21, "NY", E2:E21, "<400")<br/>
Unlike the other IF functions, this time the range you are finding the max from comes first, then the cells where your if is occurring, then your condition.

# Pivot Table Functions
Subtracting one column from another. Make sure there is ‘ ‘ around them, and a space before the ($) . Also, with this, you need to summarize by ‘Custom’.<br/>
=AVERAGE('Box Office Revenue ($)')-AVERAGE('Budget ($)')

# Product
Product will multiply cells together. =product(A2:B2)

# RIGHT/LEFT
This will return the values from the right or left of the string. <br />
=RIGHT(A2, 8)    will return the 8 most values from the right side of the cell

# Remove Duplicates
This allows you to remove duplicates from your data sheet. To use go to data > remove duplicates > select the column you want

# SORT range/sheet
Sort sheet will keep the cells together- sort range will just sort that one column without moving the other columns with it.<br />
To sort by using a formula: =SORT(A1:D6, 2, TRUE) (Range:range, column you are sorting by [it doesn’t recognize letters], TRUE for ascending order FALSE for descending) . <br /> Data range for written SORT function should never contain header row. <br /><br />

The sort from a spreadsheet's Data tab overwrites the cells containing the unsorted data with the sorted data, while a written SORT function inserts the sorted data in a different cell range.<br /><br />

To sort an entire data set by numerous conditions: highlight all of the data, click “sort range” from the data tab and then, click there is a header row, and then add your first sort, then “add another sort column” and add that.<br />

# SUMIF(S)
This one is kind of like COUNTIF. For this one, you select a column of data, and then add your condition. Then, select another column from which you are pulling a sum. Then it will take the first column, and identify only the ones that match your condition. Then, it will take the second column you identified and for all of the cells corresponding to the condition you set for the first column, it will calculate their sum. <br/>
=SUMIF(A2:A60, “=1, B2:B60)     So in this case, all of the rows where the value is 1 in the A column will find the sum of the values in column B.<br/><br/>

=SUMIF(range, criterion, sum_range)   is the basic syntax<br/><br/>

To have more than one condition, use SUMIFS. Basic syntax” <br/>
=SUMIFS(sum_range, criteria_range1, criterion1, [criteria_range2, criterion2, ...])<br/>
=SUMIFS(A2:A60, “=1”, B2:B60, C2:C60, “12/15/2020”)

# SUMPRODUCT
This allows you to multiply across rows and add down columns. So in the example of <br/>
=SUMPRODUCT(B3:B7, C3:C7) <br/>
It would be running B3*B3 + B4*C4 + …..



# TRIM
Cleans trailing and white spaces. Make a new column for your data, title it trim. Then <br />
=TRIM(A2)     …. And then apply it to the rest of the column. 


# VLOOKUP
Vertical lookup- looks up items within a column. They have to match the formatting you are looking for or it will return an error. It will not recognize column names such as A or B. VLOOKUP only returns the first match it finds, even if there are lots of possible matches. VLOOKUP can only return a value from the data to the right. It can't look left. Data analysts usually get around the problem by copying and pasting a column to the left of the data they want to look at.<br />

=vlookup(lookup_value, table_array, col_index_num, [range_lookup], true/false). <br /><br />

=VLOOKUP(103, A2:B26, 2, FALSE) where 103 is what you want to find, A2:B26 is the range of cells you are looking in, 2 is the column you are looking in, and FALSE tells it to find an exact match (true would return a close match)<br /><br />

=vlookup(A2, ‘Employee Rates’!$A$2:$B$5, 2, FALSE)<br />
Where: A2- first employee ID number and the employee hours spreadsheet.  <br />
Then: the name of the spreadsheet we want to search in, employee rates. Use ‘ ‘ and ! <br />
This is the way to reference the other spreadsheet.  <br />
A2:B5 is range. The $ locks the cell reference. <br />
, 2 Means we are looking for a match in the second column, column B for rate of pay. FALSE- exact match <br /><br />

=vlookup(A2,Sheet2!$A$2:$D$6, 4, FALSE) A2 is the ID number - 4 because the actual pay rate came from the fourth column.




# Value
This changes a text string to a numeric format which can be read by summing or vlookup:<br />
=VALUE(A2) <br />
Make sure to use trim when cleaning data <br />


