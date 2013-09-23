## Best Excel Formula Ever
### About
This repo contains the `VBA` code for a simple (one line) custom excel function that I have found to be a significant time saver in certain situations.  
### Background
A major limitation to creating formulas in excel is the fact that you (sometimes) cannot build up the sheet, column, and row references in a formula from values stored in cells. For example, say you want to link a cell to the cell `H3` in the sheet called `sheet1`.  
  
The formula for that looks like this: `=sheet1!H3` . Let's say that you wanted to create that formula by typing `sheet1` into one cell, `H` into another cell, and `3` into another cell and then somehow concatenate them together. That cannot be done in excel without a custom `VBA` function. Excel wants to concatenate the value of those three cells as a string rather than converting them into a formula referencing another cell.
  
That is the purpose of this function.
### The VBA Function
Here is the simple VBA function that accomplishes this task:

```
Function befe(col, row, sheet)
  
  befe = Range(sheet & "!" & col & row).Value
  
End Function
```
  
You can see that the function takes three arguments:  
  
1. The column of the cell to be referenced,  
2. The row of the cell to be referenced, and  
3. The sheet that contains the cell to be referenced
  
These three arguments can now reside in cells in the excel workbook and be referenced by the `=befe()` function.


### Use Cases
Let's take the example provided in the `Example.xlsm` workbook. Assume you have revenue data that is presented in the style shown in `sheet1`:    
  
[insert image]  
  
Now let's assume that you want to build a summary table of this data by year as seen below:  
  
[insert image ]
  
You can see that simply entering a formula into the 2013 column and copying over to the other years will not work as the spacing between the annual (2013, 2014, 2015, 2016, 2017) values varies throughout the sheet.
  
This is a perfect example of when the `=befe()` function can save you significant amounts of time (assuming your real world data set is large, which is usually the case). 


### How to Use the Formula
See the `.xlsx` file in the repo for an example implementation of this formula. There are many use cases, but the file shows one specific case where this is useful.  
  
The function is best implemeneted by creating a row above your table with the column values of the cells to be referenced, a column to the left of your table with the row values of the cells to be referenced, and a static cell to contain the sheet to be referenced by your table. Then the `=befe()` formula can be entered once with the appropriate `$` tags and simply copied to the entire table.
### Installation
To install the custom function you can simply copy the contents of the `VBA` module in this repo, go to Excel > Tools > Macros > Visual Basic Editor. Then insert a new module and paste the code. You can then return to your excel worbook and insert the formula into a cell as if it were a built-in formula.  
  
If this is a formula that you find yourself using frequently and therefore do not want to manually add to each excel workbook, then see this [article](http://office.microsoft.com/en-us/excel-help/copy-your-macros-to-a-personal-macro-workbook-HA102174076.aspx) about using a Personal Workbook to store custom VBA functions that will be available in every workbook automatically.