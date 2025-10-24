# 🕷️ Quotes to Scrape – Final Clean Spider  

> **A polished Python Scrapy automation that extracts, cleans, and exports quotes and author details from [Quotes to Scrape](http://quotes.toscrape.com).**  

---

### 🌟 Overview  
This scraper sequentially crawls every quote and author page, normalizes text, and generates both CSV and Excel outputs — fully cleaned, styled, and client-ready.  

**Key Features:**  
- 🕸️ Sequential page crawling  
- ✍️ Author details (DOB, birthplace, bio)  
- 🧹 Data normalization & cleaning  
- 📊 Dual export (CSV + Excel)  
- 🎨 Excel styling with professional formatting  
- 🕑 Timestamp and source metadata included  

---

### ⚙️ Workflow  
1️⃣ Crawl quote pages and follow author profiles  
2️⃣ Extract and clean all data fields  
3️⃣ Save as `quotestoscrape.csv`  
4️⃣ Reformat and style in Excel with OpenPyXL  
5️⃣ Add source and timestamp metadata for traceability  

---

### 🛠️ Tech Stack  
| Component | Tools |
|------------|-------|
| **Framework** | Scrapy |
| **Language** | Python 3.8+ |
| **Libraries** | OpenPyXL, Regex, Unicodedata |
| **Output** | CSV, Excel |

---

### 🚀 Run Locally  
```bash
pip install scrapy openpyxl
python quotes_final_clean.py
```

Outputs:  
📁 `quotestoscrape.csv`  
📗 `quotestoscrape.xlsx`

---

### 👨‍💻 Developer  
**Onyekachi Ejimofor**  
Python Developer | Web Scraping & Data Automation  

🔗 [LinkedIn](https://www.linkedin.com/in/onyekachiejimofor)  
💼 [GitHub](https://github.com/Onyeksman)  
📧 [onyeife@gmail.com](mailto:onyeife@gmail.com)

