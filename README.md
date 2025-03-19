Variance Report Automator --version 3.1.1  
The VRA is designed to automate and configure a variance report at the top level for the accounting team at Hempel Companies.  
To use, navigate to hempel-ai.com  
It functions as follows:  
1. Input, raw export for variance report uploaded via server  
2. Output, finalized export for variance report downloaded directly to computer
3. Functions within the code
   3.1. Sets print titles for rows one through seven  
   3.2  Sets narrow margins (left:0.25, right:0.25, top:0.75, bottom:0.75, header:0.3, footer:0.3)  
   3.3  Sets page width to fit with page columns  
   3.4  Deletes rows three through six  
   3.5  Unmerges cells in the first three rows, then merges cells in those respective rows from A to I  
   3.6  Merges cells in row five from B to E and from F to I  
   3.7  Center aligns cells A1, A2, and A3  
   3.8  Wraps text for all cells in column J  
   3.9  Sets the header for column J and apllies all header styles (Bold and thin bottom border)  
   3.10  Sets the width of column J to 40  
   3.11  Center aligns all cells in rows nine and below  
   3.12  Left aligns all cells in column A for rows four and below  
   3.13  Removes the cells with the value "Created on:"  
   3.14  Autofits row height for all rows  
   3.15  Formats and names the output file using valkues in cells A1 and A3.
   
Please note: You will have to autofit row height within excel.  
CTRL + A  
ALT + H + O + A  
This will autofit all rows, or feel free to use a shortcut.  
Code will only work with a month ending and a period ending columns.  
Program will not start if the uploaded document is currently opened.
