<img width="1536" height="169" alt="ChatGPT Image Sep 24, 2026, 11_24_51 AM" src="https://github.com/user-attachments/assets/8971dcd1-8a00-4fbd-b71c-b1a7f227829c" />


# PR.1 Fundamental Booster -- Excel Project

## 📊 Project Overview

**PR.1 Fundamental Booster** is an Excel practical project designed to
strengthen important Excel skills through real-world examples.

This project covers:

-   Data Input and Formatting
-   Relative and Absolute Cell References
-   IF 
-   IF with AND / OR
-   COUNTIFS, SUMIFS, AVERAGEIFS
-   VLOOKUP
-   INDEX + MATCH
-   XLOOKUP

The project uses practical datasets such as:

-   Student Grades
-   Employee Data
-   Sales Data
-   Product Prices
-   Salesperson Performance

------------------------------------------------------------------------

# 📁 Workbook Structure

The workbook can contain the following worksheets:

1.  **Students Grade** -- Student names, subjects, marks and grades
2.  **Project Instructions** -- Tasks and formulas to complete
3.  **Employee Data** -- Employee IDs, names, departments and salaries
4.  **Sales Data** -- Products, regions, salespersons, dates and sales
5.  **Lookup Data** -- Product codes, prices and employee information

------------------------------------------------------------------------

# 🎯 Main Project Tasks

1.  Apply IF formulas to classify student grades.
2.  Use COUNTIFS to count students scoring above 60.
3.  Use VLOOKUP to find product prices from sales data.
4.  Apply XLOOKUP to fetch employee salaries dynamically.
5.  Format and analyze the date of joining in Employee Data.
6.  Use FILTER to extract students with grades above 80.

------------------------------------------------------------------------

# 1. Relative & Absolute Cell References

## Relative Reference

A relative reference changes when a formula is copied.

Example:

``` excel
=B2*C2
```

If the formula is copied down, it becomes:

``` excel
=B3*C3
```

## Absolute Reference

An absolute reference remains fixed.

Example:

``` excel
=B2*$D$1
```

Here `$D$1` will not change when the formula is copied.

## Mixed References

Examples:

``` excel
=$A2
=A$2
=$A$2
```

-   `$A2` → Column A fixed, row changes
-   `A$2` → Row 2 fixed, column changes
-   `$A$2` → Both fixed

------------------------------------------------------------------------

# 2. Formatting and Data Input

Excel formatting can be used to make data clear and professional.

Common formatting:

-   Bold
-   Italic
-   Font size
-   Currency
-   Percentage
-   Date
-   Number
-   Decimal places
-   Borders
-   Alignment
-   Cell color

Example:

  Product     Price   Tax
  --------- ------- -----
  Pen           100   10%
  Book          200   10%
  Bag           500   10%

Tax calculation:

``` excel
=B2*$C$2
```

------------------------------------------------------------------------

# 3. IF Formula

The IF function checks a condition and returns one result if TRUE and
another if FALSE.

## Syntax

``` excel
=IF(logical_test,value_if_true,value_if_false)
```

## Example

``` excel
=IF(B2>=50,"Pass","Fail")
```

If marks are 70:

``` text
Pass
```

If marks are 40:

``` text
Fail
```

------------------------------------------------------------------------


# 4. IF with AND

AND requires all conditions to be TRUE.

## Example

Find students who scored above 80 in both Math and Science.

``` excel
=IF(AND(B2>80,C2>80),"Yes","No")
```

Example:

  Student     Math   Science Result
  --------- ------ --------- --------
  Ravi          85        90 Yes
  Priya         90        70 No

------------------------------------------------------------------------

# 5. IF with OR

OR returns TRUE when at least one condition is TRUE.

## Example

``` excel
=IF(OR(B2>80,C2>80),"Eligible","Not Eligible")
```

If either Math or Science is above 80, the student is eligible.

------------------------------------------------------------------------

# 6. COUNTIFS

COUNTIFS counts cells that satisfy multiple criteria.

## Syntax

``` excel
=COUNTIFS(range1,criteria1,range2,criteria2)
```

## Count students scoring above 50 in Math

``` excel
=COUNTIFS(B2:B20,">50")
```

## Count students with Math above 60 and Science above 60

``` excel
=COUNTIFS(B2:B20,">60",C2:C20,">60")
```

------------------------------------------------------------------------

# 7. SUMIFS

SUMIFS adds values that satisfy conditions.

## Syntax

``` excel
=SUMIFS(sum_range,criteria_range1,criteria1,...)
```

## Example

Suppose:

-   Column A = Region
-   Column B = Product
-   Column C = Sales

To calculate sales for the Rajkot region:

``` excel
=SUMIFS(C2:C100,A2:A100,"Rajkot")
```

To calculate sales for Rajkot and Laptop:

``` excel
=SUMIFS(C2:C100,A2:A100,"Rajkot",B2:B100,"Laptop")
```

------------------------------------------------------------------------

# 8. AVERAGEIFS

AVERAGEIFS calculates the average of values that meet conditions.

## Syntax

``` excel
=AVERAGEIFS(average_range,criteria_range,criteria)
```

## Example

Calculate average score of students scoring above 60:

``` excel
=AVERAGEIFS(B2:B20,B2:B20,">60")
```

------------------------------------------------------------------------

# 9. VLOOKUP

VLOOKUP searches for a value in the first column of a table and returns
information from another column.

## Syntax

``` excel
=VLOOKUP(lookup_value,table_array,col_index_num,FALSE)
```

## Example

  Product Code   Product     Price
  -------------- --------- -------
  P101           Pen            20
  P102           Book          100
  P103           Bag           500

Find the price of P102:

``` excel
=VLOOKUP("P102",A2:C4,3,FALSE)
```

Result:

``` text
100
```

## Important

`FALSE` means exact match.

------------------------------------------------------------------------

# 10. INDEX Function

INDEX returns a value from a specified position.

## Syntax

``` excel
=INDEX(array,row_num)
```

## Example

``` excel
=INDEX(B2:B10,3)
```

This returns the third value from B2:B10.

------------------------------------------------------------------------

# 11. MATCH Function

MATCH finds the position of a value.

## Syntax

``` excel
=MATCH(lookup_value,lookup_array,0)
```

## Example

``` excel
=MATCH("Laptop",A2:A10,0)
```

If Laptop is the 5th item, the result is:

``` text
5
```

------------------------------------------------------------------------

# 12. INDEX + MATCH

INDEX and MATCH can work together for flexible lookups.

## Formula

``` excel
=INDEX(C2:C10,MATCH(E2,A2:A10,0))
```

Meaning:

1.  MATCH finds the position of E2 in A2:A10.
2.  INDEX returns the corresponding value from C2:C10.

------------------------------------------------------------------------

# 13. TEXT Functions

TEXT functions help manipulate and format text.

## UPPER

Converts text to uppercase.

``` excel
=UPPER(A2)
```

Example:

``` text
hiteshi
```

becomes:

``` text
HITESHI
```

## LOWER

Converts text to lowercase.

``` excel
=LOWER(A2)
```

------------------------------------------------------------------------

# 14. XLOOKUP

XLOOKUP is a modern lookup function.

## Syntax

``` excel
=XLOOKUP(lookup_value,lookup_array,return_array)
```

## Example

Find employee salary:

``` excel
=XLOOKUP(E2,A2:A20,D2:D20)
```

Meaning:

-   Search E2 in A2:A20
-   Return salary from D2:D20

## With Not Found Message

``` excel
=XLOOKUP(E2,A2:A20,D2:D20,"Not Found")
```

------------------------------------------------------------------------

# 15. XMATCH

XMATCH returns the position of a value in a range.

## Syntax

``` excel
=XMATCH(lookup_value,lookup_array)
```

## Example

``` excel
=XMATCH("Laptop",A2:A20)
```

It returns the position of Laptop.

------------------------------------------------------------------------


# 16. Math Functions

## ROUND

Rounds a number to a specified number of digits.

``` excel
=ROUND(A2,2)
```

Example:

``` text
25.678 → 25.68
```


------------------------------------------------------------------------



------------------------------------------------------------------------

# 17. Complete Formula Quick Reference

  Topic                Formula
  -------------------- --------------------------------------
  Relative Reference   `=B2*C2`
  Absolute Reference   `=B2*$D$1`
  IF                   `=IF(B2>=50,"Pass","Fail")`
  AND                  `=AND(B2>80,C2>80)`
  OR                   `=OR(B2>80,C2>80)`
  IF + AND             `=IF(AND(B2>80,C2>80),"Yes","No")`
  IF + OR              `=IF(OR(B2>80,C2>80),"Yes","No")`
  COUNTIFS             `=COUNTIFS(B2:B20,">50")`
  SUMIFS               `=SUMIFS(C2:C100,A2:A100,"Rajkot")`
  AVERAGEIFS           `=AVERAGEIFS(B2:B20,B2:B20,">60")`
  VLOOKUP              `=VLOOKUP(E2,A2:C20,3,FALSE)`
  INDEX                `=INDEX(B2:B20,3)`
  MATCH                `=MATCH(E2,A2:A20,0)`
  INDEX + MATCH        `=INDEX(C2:C20,MATCH(E2,A2:A20,0))`
  UPPER                `=UPPER(A2)`
  LOWER                `=LOWER(A2)`
  XLOOKUP              `=XLOOKUP(E2,A2:A20,D2:D20)`
  XMATCH               `=XMATCH(E2,A2:A20)`
  ROUND                `=ROUND(A2,2)`

------------------------------------------------------------------------

# 🧪 Practice Tasks

## Student Grade Practice

Create this table:

  Student     Math   Science Grade
  --------- ------ --------- -------
  Ravi          85        90 
  Priya         65        72 
  Jay           45        55 
  Neha          92        88 

Tasks:

1.  Calculate Grade using Nested IF.
2.  Find students scoring above 60.
3.  Find students scoring above 80 in both subjects.
4.  Use FILTER to display students with marks above 80.

------------------------------------------------------------------------

## Sales Practice

Create:

  Region      Product      Sales
  ----------- ---------- -------
  Rajkot      Laptop       50000
  Ahmedabad   Mouse         5000
  Rajkot      Laptop       45000
  Surat       Keyboard      7000

Tasks:

1.  Calculate total Rajkot sales using SUMIFS.
2.  Calculate Rajkot Laptop sales using SUMIFS.
3.  Count sales above 10000.
4.  Use lookup functions to find product information.

------------------------------------------------------------------------

## Employee Practice

Create:

  ID     Name    Department     Salary Joining Date
  ------ ------- ------------ -------- --------------
  E101   Ravi    IT              30000 01/01/2024
  E102   Priya   HR              35000 15/06/2023
  E103   Jay     Sales           28000 10/02/2022

Tasks:

1.  Find salary using XLOOKUP.
2.  Find employee details using INDEX + MATCH.
3.  Calculate years of service.
4.  Extract joining year using YEAR.
5.  Extract joining month using MONTH.

------------------------------------------------------------------------

# ✅ Project Completion Checklist

-   [ ] Data entered correctly
-   [ ] Data formatted professionally
-   [ ] Relative references practiced
-   [ ] Absolute references practiced
-   [ ] IF completed
-   [ ] Nested IF completed
-   [ ] AND / OR completed
-   [ ] COUNTIFS completed
-   [ ] SUMIFS completed
-   [ ] AVERAGEIFS completed
-   [ ] VLOOKUP completed
-   [ ] INDEX + MATCH completed
-   [ ] TEXT functions completed
-   [ ] XLOOKUP completed
-   [ ] XMATCH completed
-   [ ] INDIRECT completed
-   [ ] OFFSET completed
-   [ ] Date functions completed
-   [ ] Math functions completed
-   [ ] FILTER completed
-   [ ] Final workbook checked

------------------------------------------------------------------------

# 📌 Learning Outcome

After completing this project, you should be able to:

-   Enter and format Excel data.
-   Use relative, absolute and mixed references.
-   Create logical formulas.
-   Analyze student and sales data.
-   Search data using VLOOKUP and XLOOKUP.
-   Combine INDEX and MATCH.
-   Manipulate text using Excel TEXT functions.
-   Work with dates and times.
-   Create dynamic references and ranges.
-   Extract filtered results using FILTER.
-   Build practical Excel worksheets for academic and business use.

------------------------------------------------------------------------

# 🏁 Final Project Goal

The main goal of **PR.1 Fundamental Booster** is to move from basic
Excel formulas to practical data analysis.

Recommended learning order:

**Data Input → Formatting → References → IF → AND/OR →
COUNTIFS/SUMIFS/AVERAGEIFS → VLOOKUP → INDEX/MATCH → TEXT → XLOOKUP →
XMATCH → INDIRECT → OFFSET → Date/Time → Math → FILTER**

------------------------------------------------------------------------

## 👩‍💻 Author

**Project:** PR.1 Fundamental Booster\
**Tool:** Microsoft Excel\
**Level:** Beginner to Intermediate\
**Purpose:** Excel Practical Learning & Data Analysis
