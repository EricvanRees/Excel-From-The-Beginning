# Notes for "Excel from the Beginning" course

## 1. Notes for Payroll project

Spread sheet cells have rows and numbers

* How to Create a Formula:

1) Click to select a cell, then enter "=" to start a formula
2) Next, enter the formula. For example "=C4*D4"
3) Press enter to execute the formula

* How to copy a formula
  1) Use "copy" and "paste" to copy a formula to a single cell
  2) Fill down values using a formula by clicking the green square in a cell and dragging down the selected cell

* =Max(range)
Finds maximum value in a range

* =Avg(range)
Calculates average from a range of values
"Decrease decimals" = "Minder decimalen"

* =Min(range)
Finds minimum value in range

* Use logical test
  click a cell, type "=if(D4>40,D4-40,0)"
  This means: if the value of D4 is bigger than 40, calculate overtime as the value of D4 minus 40, else value is zero.

  NL: =ALS(D4>40;D4-40;0)

* Relative vs. absolute referencing
  
  ="C4*D4" vs "$C4*D4"
  "$C4" tells Excel to always reference column C 

  In Excel, relative references change as you copy a formula, while absolute references remain constant regardless of where you copy the formula, denoted by a dollar sign ($) before the column and/or row. 


  **How To Create a Header Row in Excel Using 3 Methods** [link](https://www.indeed.com/career-advice/career-development/how-to-create-header-row-in-excel) 

## 2. Notes for Gradebook project

Use conditional formatting with "stijlen -> voorwaardelijke opmaak -> Pictogramseries"

These have to be created by column, you cannot apply them for multiple rows and columns.

**Conditional formatting**
Only format cells with values less than 50%:

Conditional formatting -> 
"markeringsregels voor cellen -> kleiner dan"

**OR formula**
Combine a series of logical questions:
Is the score for each field less than 50%?
"=OF(H4<0.5;I4<0.5;J4<0.5;K4<0.5)"

**Insert Bar Chart**

Invoegen -> kolom- of staafdiagram

Change data labels-> right-click the graph, choose "Gegevens selecteren". 

## 3. Notes for Decision chart project

Show top 10% fields of a column:

"Stijlen" -> "voorwaardelijke opmaak" -> "Regels voor bovenste/onderste" -> "bovenste 10%"

## 4. Notes for Sales items project

Formulas used in this lesson:
1) Text to columns - split single column into separate columns (first name - last name)
2) If 
3) Sumif - pick items together based on self-defined criteria
4) Sort - database command
5) Filter - database command
6) Pivot tables - summary of each sales person's sales
7) Pie chart - charting the results

**Text wrap function**
Text wrap = "uitlijning" -> "terugloop"

**Text to columns**
Select row of choice, next click "Data" -> "Text to columns"

Next, select delimiter (= divider) to split the column.

**Quickly fill an entire row with data**
Fill in the formula(s) to apply, then manually select all rows, next select "bewerken -> doorvoeren -> omlaag".

**sumif**
1) sum of items valued at more than $50
"sumif" = "SOM.ALS" = SUM(range, critera) = sumif(f2:f172,">50") - This will sum items only if they're greater than 50.

2) sum of items valued at $50 or less
  "sumif(f2:f172, "<=50)"

**sorting & filtering**

1) sorting data
First, select all rows and columns. Next, click:
"Gegevens" -> "Sorteren en filteren", select column to sort on (letters or numbers).

2) filtering data
By clicking filter, all column headers will have a drop-down button. You can now select or filter on individual values.

**Pivot table**
First, select all rows and columns. Then, from the Insert tab, choose "Pivot table" ("Invoegen" -> "Draaitabel"). Click "OK".

A new sheet opens, select the columns you want to include. Change the values in the columns by clicking the value field of "waarden" (rechtsonder in beeld) en dan "waardenveldinstellingen". 

**Pie chart**
Click "insert" -> "Pie chart" ("Invoegen" -> "Cirkel of ringdiagram invoegen"). Right-click the pie chart and choose "add data labels" for adding data to the pie chart.












 
