Sure! Here are the formulas for **Date Functions** and **Text Functions** in Excel:

### **Date Functions:**  
1. **Current date and time:** `=NOW()`  
2. **Extract year:** `=YEAR(A1)`  
3. **Days between two dates:** `=A2 - A1`  
4. **Weekday name:** `=TEXT(A1, "dddd")`  
5. **Add 30 days to a date:** `=A1 + 30`  
6. **Calculate age from birthdate:** `=DATEDIF(A1, TODAY(), "Y")`  
7. **Extract month name:** `=TEXT(A1, "mmmm")`  
8. **Last day of the month:** `=EOMONTH(A1, 0)`  

### **Text Functions:**  
1. **Concatenate with space:** `=A1 & " " & B1` or `=TEXTJOIN(" ", TRUE, A1, B1)`  
2. **Extract first name:** `=LEFT(A1, FIND(" ", A1) - 1)`  
3. **Extract last 4 characters:** `=RIGHT(A1, 4)`  
4. **Convert to uppercase:** `=UPPER(A1)`  
5. **Extract domain from email:** `=RIGHT(A1, LEN(A1) - FIND("@", A1))`  
6. **Remove extra spaces:** `=TRIM(A1)`  
7. **Count characters (including spaces):** `=LEN(A1)`  
8. **Replace "banana" with "grape":** `=SUBSTITUTE(A1, "banana", "grape")`  

Let me know if you need any modifications! 🚀
