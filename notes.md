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


 
