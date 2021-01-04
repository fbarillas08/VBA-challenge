# VBA-challenge
Multi_Year Stock Price Analysis

This is a Visual Basics Programming Project.  The objective was to summarize the annual performance of
over 12,000 stocks during three consecutive years: 2014, 2015, and 2016.

## The analytical process is as follows:
1. Extract each unique ticker symbol using an if statement that compares each row to its next to
   identify when the ticker symbol changes.
2. Once there is a change identify, the End-of-Year (EOY) Price for the current ticker symbol is stored
   along with the Beginning-of-Year (BOY) Price for the next ticker symbol.   Those two values are
   populated into the Summary Table for each stock.
3. While there is no change in ticker name, the volume of trade for each stock is accumulated in a
   variable called tickervolumetotal.  At each change the accumulated value is populated in the
   Summary Table
4. Once the last row is verified to be a blank, the program sets to compute the delta between the
   BOY and EOY prices of each stock and the % annual change in price.

## The formatting process is as follows
 Headers are formatted for easy reading using the Range function

 Conditional formatting for cells is done to the column recording the Annual % Price Change for each stock.
 Three color variations are exhibited at the discretion of the programmer as follows:
 - Values 15% or larger are green
 - Value 0% up to 15% not inclusively are yellow
 - Negative Values are redish

 In addition, all values in the Summary Table and the Bonus Table have been formatted either as
 two-decimal numbers or a two-decimal percentage.  The font format for the delta change in price
 includes a provision to change the font color to red for negative numbers.

 ## BONUS Table Approach

 To find the outliers (Maximum Volume Traded, Largest % Increase, and Lowest % Increase) three variables
 were defined: maxgrowth, mingrowth, and maxvolume.  Application.WorksheetFunctions for each of these
 computations are used on the appropriate column

 The values of the three variables are then populated into the Bonus Table.

 An iterative loop is done to review the ticker column and retrieve the ticker symbol whose annual growt
 or volume matches the computed variables above.

 Once a match is made, the ticker symbol is populated in the corresponding bonus row.

 Once all computations and formating are completed, a text box message notifies the operator.

 ## For questions or clarifications please contact me at:
 ## fbarills08@gmail.com
 ## fbarillas@vectoraldata.com


