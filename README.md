# 🧠 Streamlined Fuzzy Lookup ETL Pipeline  

## 📖 Overview  
This project automates an **end-to-end data reconciliation workflow** using **Python**, **pandas**, and **RapidFuzz**.  
It connects directly to **Microsoft SQL Server**, cleans and normalizes address data, performs **fuzzy address matching** against ROI tables, and outputs tiered Excel reports ranked by similarity scores.  

This repository serves as a **demonstration of professional ETL automation** and data quality management using fuzzy logic and efficient vectorized computation.  

---

## ⚙️ Key Features  
- 🔌 **Direct SQL Server Integration** – Connects using `pyodbc` and `SQLAlchemy` for automated data extraction.  
- 🧹 **Data Cleansing and Normalization** – Cleans inconsistent addresses, ZIP codes, and region mappings using robust validation logic.  
- 🧩 **Fuzzy Matching Engine** – Uses `RapidFuzz` for address similarity scoring with adjustable thresholds (≥95%, 90–94%, 85–89%).  
- 📊 **Automated Excel Reporting** – Exports three categorized sheets (`Match_95_and_Above`, `Match_90_to_94`, `Match_85_to_89`) in a single Excel workbook.  
- 🔍 **Difference Classification** – Highlights differences between strings and classifies errors as *1 Letter Off*, *2 Letters Off*, or *Numerical*.  
- 🗺️ **Regional Assignment Logic** – Automatically inserts regional ownership data based on subregion mapping.  
- ⏱️ **Performance Optimized** – Replaces a multi-hour manual reconciliation process with a single automated Python execution (~minutes).  

---

## 🧠 Technical Stack  
| Category | Tools / Libraries |
|-----------|--------------------|
| Language | Python 3 |
| Database | Microsoft SQL Server |
| Libraries | pandas, RapidFuzz, openpyxl, SQLAlchemy, pyodbc, re, difflib |
| Output | Excel (multi-sheet), CSV |
| OS Tested | Windows 10+ |

---

## 🧾 Process Flow  
1. **Extract:**  
   - Connects to SQL Server via trusted connection  
   - Executes a `.sql` query from `/query/Query.sql`  
   - Loads results into a pandas DataFrame  

2. **Transform:**  
   - Cleans and reformats address and ZIP columns  
   - Adds “Regional Assignment” column from a subregion-to-owner mapping  
   - Merges data with ROI tables for address enrichment and SHCODE lookup  

3. **Fuzzy Match:**  
   - Compares extracted addresses with ROI table addresses using RapidFuzz similarity ratios  
   - Validates ZIP code similarity (>85%)  
   - Classifies differences using `difflib.Differ`  

4. **Load (Export):**  
   - Saves results into an Excel workbook named `FuzzyMMDDYYYY.xlsx`  
   - Organizes matches into three confidence tiers  

---

## 📈 Example Output
| Sheet Name | Match Range | Description |
|-------------|-------------|-------------|
| `Match_95_and_Above` | ≥ 95% | High-confidence address matches |
| `Match_90_to_94` | 90–94% | Moderate confidence matches |
| `Match_85_to_89` | 85–89% | Low confidence matches requiring manual review |

Each record includes:
- Matched address  
- ZIP validation result  
- Difference classification (1 Letter Off, 2 Letters Off, etc.)  
- Regional assignment and Outreach ID  

---

## 📈 Outcomes  
- Decreased reporting cycle from **~2 hours to under 10 minutes**.  
- Reduced error rates and improved data integrity through automated validation.  
- Provided a reusable Python framework adaptable for other matching and ETL use cases.  

---

## 🧠 Learning Outcomes  
Through this project, I demonstrated:
- Building a complete **ETL pipeline** in Python  
- Applying **fuzzy logic algorithms** for data reconciliation  
- Using **SQLAlchemy** and `pyodbc` for enterprise-grade data access  
- Designing a scalable export process for large data sets  
- Writing clear, modular code with reproducible results  

---

## 📬 Contact  
**Kevin Braman**  
📧 [kevinbraman92@gmail.com](mailto:kevinbraman92@gmail.com)  
💼 [LinkedIn](https://www.linkedin.com/in/kevin-braman-a7974a129/)  
💻 [GitHub](https://github.com/kevinbraman92)

---

⭐ *If this project inspires you or helps your workflow, consider giving it a star!*


