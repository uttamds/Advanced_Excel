Sure! Here are **MS Excel examples** for the topic **"Data Cleaning and Preparation"**, commonly used in training sessions, especially in data analysis or business reporting contexts.

---

### 🧹 **1. Removing Duplicates**
**Scenario:** You have a customer list and some names appear more than once.  
**Example:**  
| Name     | Email               |
|----------|---------------------|
| John     | john@email.com      |
| Sarah    | sarah@email.com     |
| John     | john@email.com      |

**Action:**  
Go to **Data** → **Remove Duplicates**  
👉 Select the columns to check duplicates in.  
**Result:** Only unique entries remain.

---

### ✂️ **2. Trimming Extra Spaces**
**Scenario:** Extra spaces before/after names cause inconsistency.  
**Example (raw data):** `"  John "`  
**Action:**  
Use the formula `=TRIM(A2)` to clean the name.  
**Result:** `"John"`

---

### 🔍 **3. Finding and Replacing Text**
**Scenario:** You want to replace all instances of "N/A" with a blank.  
**Action:**  
Use **Find & Replace**:  
- Find what: `N/A`  
- Replace with: *(leave blank)*

---

### 🧽 **4. Converting Text to Proper Case**
**Scenario:** Names are all lowercase.  
**Example:** `"john smith"`  
**Formula:** `=PROPER(A2)`  
**Result:** `"John Smith"`

---

### 🧱 **5. Splitting Data into Columns (Text to Columns)**
**Scenario:** You have full names in one column and want first/last names separated.  
**Example:** `"John Smith"`  
**Action:**  
Select the column → Go to **Data** → **Text to Columns** → Choose **Delimited** → Select **Space** as delimiter.  
**Result:**
| First Name | Last Name |
|------------|-----------|
| John       | Smith     |

---

### 🔄 **6. Converting Numbers Stored as Text**
**Scenario:** Numbers are stored as text, affecting calculations.  
**Example:** `"100"` stored as text  
**Fix:**  
Use `=VALUE(A2)` or use the **Error Smart Tag** to convert to number.

---

### 🧹 **7. Dealing with Blanks**
**Scenario:** Blank cells affect formulas or charts.  
**Fix options:**  
- Use filters to find and delete them  
- Fill blanks using `=IF(A2="", "Unknown", A2)`

# Std Deviation

![image](https://github.com/user-attachments/assets/5337c1d1-f194-4d45-b0ad-9c68783b2269)

