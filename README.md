# ⚙️ Automating Data Input Workflow to Improve Accuracy & Efficiency  
**Tools:** Microsoft Excel | VBA (Macro) | Process Automation | Data Quality Control  

---

## 🧩 Project Overview  
This project was developed to address inefficiencies and frequent errors in the manual data entry process for item code creation within the company’s internal system.  

Previously, each team member had to manually fill in multiple Excel templates and send them to the IT department for code generation. This process often resulted in missing or inconsistent information, taking over **20 minutes per case** and causing operational delays.  

The goal of this project was to automate and standardize the entire data entry process, ensuring **100% data accuracy** while significantly reducing time spent on repetitive manual tasks.

---

## 💡 Business Problem  
- Input data for new item codes was scattered across multiple files, leading to incomplete or inconsistent submissions.  
- Manual filling of Excel templates was time-consuming and error-prone.  
- Each member used different file paths and templates, making standardization difficult.  

---

## 🎯 Objectives  
1. Centralize and structure all required information for item creation in a single master file.  
2. Automate the population of multiple Excel templates with pre-validated data.  
3. Minimize data entry time and eliminate human errors.  
4. Provide an easy-to-deploy solution usable by all team members without technical expertise.  

---

## ⚙️ Solution Overview  

### Step 1 — Centralized Data Collection  
- Created a **master Excel file** containing structured sheets for different item creation scenarios.  
- Added configurable input cells to store template file paths (so each user can adjust their local environment easily).  

### Step 2 — Automation with VBA  
- Built a **VBA-based automation script** that:  
  - Opens relevant templates automatically.  
  - Populates all required fields with validated master data.  
  - Creates a **timestamped folder** (date + time) to store the completed templates for record-keeping.  
  - Handles errors gracefully (using `On Error GoTo`) to ensure smooth execution and user confidence.  

### Step 3 — Deployment & Team Adoption  
- Delivered the macro-enabled Excel file (`.xlsm`) to all team members.  
- Conducted testing and setup guidance to enable VBA macros.  
- Trained users to operate the system with minimal effort (no coding knowledge required).  

---

## 📈 Results & Impact  
- ⏱ Reduced data entry time from **20+ minutes to under 2 minutes per case.**  
- ✅ Achieved **100% accuracy** in all submitted templates.  
- 👥 Deployed to the entire team (8 users), saving an estimated **80+ work hours per month.**  
- 🧠 Improved overall data governance and process consistency for IT code creation requests.  

---

## 🧠 Key Insights  
- Even widely-used tools like Excel can become **powerful automation platforms** when paired with simple VBA logic.  
- Automating repetitive business processes not only saves time but also **builds data reliability and confidence** across departments.  
- Process automation should always be paired with **user training** and clear documentation to ensure adoption and sustainability.  

---

## 🧰 Skills & Tools Used  
- **Microsoft Excel** — data structuring and user interface design  
- **VBA / Macros** — automation of repetitive workflows  
- **Process Optimization** — streamlining manual business tasks  
- **Data Quality Control** — validation and accuracy checks  
- **Documentation & User Training** — enabling non-technical adoption  

---

## 📸 Preview  
![Folder List.PNG]

