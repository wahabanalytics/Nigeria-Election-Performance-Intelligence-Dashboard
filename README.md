# 🇳🇬 Nigeria Election Performance Intelligence Dashboard

## 📊 Project Overview

The **Nigeria Election Performance Intelligence Dashboard** is an interactive data analytics project designed to analyze and visualize election performance across Nigeria.

The dashboard transforms election data into meaningful insights, allowing users to explore voter participation, election outcomes, candidate performance, political party performance, and geographical trends across different election years and locations.

The project was developed using **Microsoft Excel and VBA**, combining advanced Excel analytics with automation to create a fully interactive multi-page dashboard.

> ⚠️ **Disclaimer:** The dataset used in this project is synthetic and was created strictly for analytical and portfolio purposes. Any candidates, winners, parties, or election results displayed should not be interpreted as official historical Nigerian election records.

---

# 🎯 Business Problem

Election data can contain large volumes of information across multiple years, election types, political parties, candidates, states, and Local Government Areas.

Analyzing this information manually makes it difficult to quickly answer important questions such as:

- Which election recorded the highest voter participation?
- Which states had the highest number of accredited voters?
- Who won a particular election?
- Which political parties performed best?
- How did candidate performance vary across elections?
- How does voter turnout differ across States and LGAs?
- How do election metrics change across different election years?

This project was designed to transform raw election data into an interactive analytical solution that makes these insights easier to explore and understand.

---

# 🎯 Project Objectives

The major objectives of this project were to:

- Analyze election performance across multiple election years.
- Monitor voter participation and turnout.
- Compare registered voters, accredited voters, valid votes, and rejected votes.
- Identify election winners and winning political parties.
- Analyze candidate performance.
- Evaluate political party performance.
- Analyze election results across States and LGAs.
- Provide geographical visualization of election participation across Nigeria.
- Build an interactive dashboard with dynamic filtering and navigation.
- Create a user-friendly analytical experience using Microsoft Excel.

---

# 🛠️ Tools & Technologies Used

### 📊 Microsoft Excel

Excel was used for:

- Data Cleaning
- Data Transformation
- Pivot Tables
- Pivot Charts
- Advanced Formulas
- Named Ranges
- Helper Cells
- Form Control Combo Boxes
- Dashboard Design

### ⚙️ VBA (Visual Basic for Applications)

VBA was used to:

- Automate dashboard filtering.
- Connect Combo Box selections to slicers.
- Manage Election Year filtering.
- Manage Election Type filtering.
- Manage Zone filtering.
- Manage State filtering.
- Manage LGA filtering.
- Manage Party filtering.
- Create dependent State → LGA filtering.
- Improve dashboard response speed.
- Update winner images dynamically.
- Control interactive dashboard behavior.

---

# 📂 Dataset Overview

The dashboard analyzes synthetic election data covering:

| Category | Coverage |
|---|---|
| 📅 Election Years | 2007 – 2023 |
| 🗳️ Election Types | Presidential, Governorship, Senatorial, House of Representatives |
| 📍 Geographical Coverage | 36 States + FCT |
| 🏘️ Local Government Areas | 774 LGAs |

The dataset includes information such as:

- Election Year
- Election Type
- Zone
- State
- Local Government Area
- Political Party
- Candidate
- Registered Voters
- Accredited Voters
- Valid Votes
- Rejected Votes
- Total Votes
- Winning Candidate
- Winning Party
- Vote Share
- Winning Margin

---

# 🖥️ Dashboard Pages

## 🏠 Cover Page

The entry point into the dashboard, providing navigation into the analytical pages.
<img width="1600" height="614" alt="image" src="https://github.com/user-attachments/assets/103c38c5-de15-4b1a-a247-c649613e1f5c" />

---

## 📊 Executive Overview

Provides a high-level summary of election performance.

### Key KPIs

- 👥 Registered Voters
- 🗳️ Accredited Voters
- 📈 Overall Turnout
- ✅ Valid Votes
- ❌ Rejected Votes
- 📉 Rejection Rate

### Key Features

- Executive Insights
- Election performance trends
- State-level geographical visualization
- Dynamic Presidential Winner Card
- Interactive filtering
<img width="1418" height="697" alt="image" src="https://github.com/user-attachments/assets/f532694c-5050-420e-ba19-390cde37047d" />

---

## 📈 Election Analysis

Provides deeper analysis of election performance across:

- Election Years
- Election Types
- Candidates
- Political Parties
<img width="1410" height="697" alt="image" src="https://github.com/user-attachments/assets/8ac57aa7-c587-4f7e-b5f3-283e5d31adb7" />

---

## 🗺️ State & LGA Analysis

Provides geographical analysis of election performance at:

- State Level
- Local Government Area Level

### Key Features

- State-level analysis
- LGA-level analysis
- State vs National Turnout comparison
- Dynamic Winner Analysis
- Dependent State → LGA filtering
<img width="1424" height="696" alt="image" src="https://github.com/user-attachments/assets/5f5281c0-2148-4e08-bccc-19ec3a0a3c2d" />

---

## 👤 Candidate Analysis

Provides insights into:

- Candidate performance
- Candidate votes
- Candidate vote share
- Candidate rankings
<img width="1488" height="692" alt="image" src="https://github.com/user-attachments/assets/68ebca70-7b14-4905-b543-505c999b216b" />

---

## 🚩 Party Analysis

Provides insights into:

- Party performance
- Party vote share
- Party rankings
- Winning party performance
<img width="1451" height="689" alt="image" src="https://github.com/user-attachments/assets/84572795-dfb8-44b8-b352-9175deb06c59" />

---

## 📖 Data Dictionary

Provides explanations for important dataset fields and metrics.
<img width="1600" height="668" alt="image" src="https://github.com/user-attachments/assets/ecb8c624-21b6-413e-bf49-305ac26c9385" />

---

## ℹ️ About Project

Provides information about:

- Project purpose
- Dataset coverage
- Dashboard functionality
- Navigation instructions
<img width="1600" height="623" alt="image" src="https://github.com/user-attachments/assets/f1a3740c-07eb-4f61-871f-f38ffe27e395" />

---

# 🎛️ Interactive Filtering System

The dashboard uses six major filters:

- 📅 Election Year
- 🗳️ Election Type
- 🌍 Zone
- 📍 State
- 🏘️ LGA
- 🚩 Party

The filtering system uses **Form Control Combo Boxes connected to VBA-powered slicers**.

This creates a cleaner dashboard interface while maintaining powerful interactive filtering.

---

# 🔗 Dependent State → LGA Filtering

One of the key interactive features implemented in this project is dependent filtering.

When a user selects a **State**, the LGA filter automatically updates to display only LGAs belonging to the selected State.

This improves usability and prevents invalid geographical selections.

---

# ⚡ Performance Optimization

One of the major technical challenges encountered during development was slow LGA filtering.

Initially, selecting an LGA could take several minutes to respond.

The VBA filtering process was optimized using:

- `PivotTable.ManualUpdate`
- Controlled calculation settings
- Disabled Screen Updating during processing
- Disabled Events during processing
- Reduced unnecessary PivotTable refreshes
- Faster slicer filtering using `VisibleSlicerItemsList`

### 🚀 Result

The LGA filtering response improved significantly and became almost instantaneous.

---

# 🖼️ Dynamic Winner Cards

The dashboard includes dynamic Winner Cards that display relevant election information based on selected filters.

The cards provide information such as:

- Winner
- Winning Party
- Winning Votes
- Vote Share
- Winning Margin
- Dynamic Winner Images

VBA automation was used to dynamically update winner images while preserving the user's current dashboard page.

---

# 🧠 Dynamic Executive Insights

The Executive Overview includes automatically generated insights based on election data and user selections.

The insights dynamically communicate:

- Election winners.
- Voter participation.
- Highest accredited voter participation.
- Winning margins.
- Election performance trends.

The insight system uses:

- `INDEX`
- `MATCH`
- `MAX`
- `IFERROR`
- PivotTable outputs
- Helper Cells

---

# 🗺️ Geographical Analysis

The dashboard includes an interactive geographical visualization of Nigeria.

Users can explore how election participation varies across States using the dashboard filters.

---

# ⚠️ Technical Challenges & Solutions

## 🔴 Challenge 1: Slow LGA Filtering

**Problem:**  
LGA filtering initially took several minutes to respond.

**Solution:**  
The VBA filtering process was optimized by reducing unnecessary refresh operations and controlling PivotTable updates.

**Result:**  
⚡ LGA filtering became significantly faster.

---

## 🔴 Challenge 2: Dashboard Redirection

**Problem:**  
Changing filters on certain pages redirected users to the Executive Overview.

**Solution:**  
The VBA macro was updated to preserve the user's original worksheet before updating winner images and restore the sheet afterward.

**Result:**  
Users remain on their current dashboard page while filtering.

---

## 🔴 Challenge 3: Default Winner Results

**Problem:**  
Winner Cards displayed default election results when required filters had not been selected.

**Solution:**  
Conditional formulas were implemented to ensure cards remain blank until valid filter conditions are met.

---

## 🔴 Challenge 4: State vs National Turnout Evaluating Too Early

**Problem:**  
The visualization responded when only a State was selected.

**Solution:**  
Additional conditional logic was added so the analysis only evaluates after the required Election Year and Election Type selections are made.

---

## 🔴 Challenge 5: PivotTable Overwrite Warning

**Problem:**  
Certain filter selections triggered a PivotTable overwrite warning.

**Solution:**  
The PivotTable layout and refresh behavior were corrected.

---

# 🎨 Dashboard Design Approach

The dashboard was designed around the following principles:

### ✅ User-Friendly Navigation

Users can move easily between dashboard pages.

### ✅ Consistent Design

A consistent visual identity was maintained throughout the dashboard.

### ✅ Interactive Analysis

Users can dynamically explore election data using filters.

### ✅ Executive Reporting

The Executive Overview provides important information at a glance.

### ✅ Minimal Scrolling

The dashboard was designed as a polished analytical interface rather than a long scrolling report.

---

# 🧠 Skills Demonstrated

## 📊 Data Analysis

- Data Cleaning
- Data Transformation
- KPI Development
- Trend Analysis
- Comparative Analysis
- Geographical Analysis

## 📈 Data Visualization

- Dashboard Development
- Chart Selection
- Analytical Storytelling
- Executive Reporting

## ⚙️ Microsoft Excel

- Pivot Tables
- Pivot Charts
- Advanced Formulas
- Named Ranges
- Helper Cells
- Form Controls

## 💻 VBA

- Automation
- Slicer Management
- Dynamic Filtering
- PivotTable Optimization
- Shape Manipulation
- Dynamic Image Updates
- Performance Optimization

---

# 📁 Project Structure

Nigeria-Election-Performance-Intelligence-Dashboard/
│
├── 📁 Dashboard
│   └── Nigeria Election Performance Intelligence Dashboard.xlsm
│
├── 📁 Dataset
│   └── Nigeria Election Dataset.xlsx
│
├── 📁 Documentation
│   └── Nigeria Election Performance Intelligence Dashboard Project Report.pdf
│
├── 📁 Screenshots
│   ├── Cover Page.png
│   ├── Executive Overview.png
│   ├── Election Analysis.png
│   ├── State & LGA Analysis.png
│   ├── Candidate Analysis.png
│   ├── Party Analysis.png
│   └── About Project.png
│
└── README.md

# 🚀 How to Use the Dashboard

Follow these steps to explore the dashboard:

1. Download the Excel dashboard file.
2. Open the file using **Microsoft Excel Desktop**.
3. Enable **Macros** when prompted.
4. Use the navigation buttons to move between dashboard pages.
5. Use the interactive **Combo Box filters** to explore the data.
6. Select a **State** to view the corresponding LGAs.
7. Explore election performance across different:
   - Election Years
   - Election Types
   - Geographical Locations
   - Candidates
   - Political Parties

---

# 🔮 Future Improvements

Potential future improvements for this project include:

- 🔄 Rebuilding the dashboard in **Power BI**
- ⚙️ Using **Power Query** for automated data transformation
- 🗂️ Creating a relational data model
- 📊 Adding **DAX measures**
- 🔍 Implementing drill-through analysis
- 🌐 Connecting to live or regularly updated data sources
- 🤖 Adding predictive analytics and advanced analytical capabilities

---

# ⚠️ Dataset Disclaimer

This project uses a **synthetic dataset created strictly for learning and portfolio purposes**.

The election results, candidates, winners, political parties, and other information displayed in the dashboard should **not be interpreted as official Nigerian election records**.

This project was developed to demonstrate practical skills in:

> **Data Analysis | Dashboard Development | Data Visualization | Excel Automation | VBA | Interactive Reporting**

---

# 👨‍💻 Author

## **Wahab Sodiq**

### **Aspiring Data Analyst | Excel & Dashboard Development**

---

## 🛠️ Skills Applied

- 📊 **Microsoft Excel**
- 📈 **Data Analysis**
- 📉 **Data Visualization**
- ⚙️ **VBA Automation**
- 🎨 **Dashboard Development**

---

⭐ **If you found this project interesting, feel free to explore the repository and check out the dashboard!**
