# Krutika-Desai---Data-Viz
This is the link for my GitHub repository  https://github.com/desaikx2/Krutika-Desai---Data-Viz
In this link you will find a folder named "Module 2 VB" 
This folder contains another folder named "Module 2 Challenge VB" - this contains (4) Screenshots one for each quarter showing results and the excel workbook with the VB Script. (https://github.com/desaikx2/Krutika-Desai---Data-Viz/tree/main/Module%202%20VB/Module%202%20Challenge%20VB) 
The Seperate VBA Script is store as "QuaterlyStockSummary.vbs" in folder "Module 2 Challenge VB"
Getting Tickr Symbo:
I used the logic to use For Loop for all the rows, under this I am checking the cells with same value for first and following cell and the place where this does not match - the code will read that as end of quarter and bring that as closed price and ticker symbol in new table. This value is stored in a variable called  FindStickerRowCount

Quarterly Change from the opening price at the beginning of a given Qtr to the closing price at the end of that Qtr.
Based on the cell where the symbol differs will give me the closed price. Open price will be where the For loop starts and Qtr change will be Close - Open.

%Change was obtained by (Close-Open)/Start and used Number format to give percent.

Total Stock Volume = i gave me the last cell where the Tickr symbol differed which gave me the row count, the first cell was already stored using these two I applied SUM for all the cells between these two inclusive of each mentioned.

Conditional formatting used to show positive and negative quartely change values.

Greatest % Increase - variable PreGreatPercent was created and this was used to compare next value, if > than previous than it saves the value in same for loop, this will finally give the largest value back.
Greatest % Decrease - same approach as above was used

