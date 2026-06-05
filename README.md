# Goal Setting Tool

A Streamlit-based enterprise application for setting national sales goals and allocating them across territories using data-driven allocation models. Built for sales operations teams that need a fair, transparent, and repeatable way to set territory-level goals.

---

## Overview

The Goal Setting Tool replaces manual spreadsheet-driven goal-setting with a structured, visual workflow. Users upload historical sales data, validate and clean it through automated checks, set a national sales goal, and distribute it across territories using one of three allocation models. A Back Testing module validates the chosen model against past quarters before the goals are finalised.

The final territory goals can be reviewed in interactive tables and charts, and downloaded as a professionally formatted Excel file.

---

## Key Features

### 📁 Input & Validation
- Upload sales data in CSV or Excel format
- Automatic schema validation against required columns
- Data quality checks for missing values, duplicate rows, and statistical outliers (±3σ)
- Date range filtering and trend analysis (weekly, monthly, quarterly views)
- Cleaned data can be downloaded as a formatted Excel file

### 🎯 National Goal Setting
Three allocation models to distribute the national goal across territories:

- **Weighted Model** — Blends recent and prior period sales using configurable weights, with volume smoothing, growth cap/floor controls, and a two-pass redistribution mechanism. Includes a pro-rata fallback to guarantee the total allocation always matches the National Goal.
- **Fair Share Model** — Distributes the goal in proportion to each territory's recent sales contribution.
- **Equal Allocation Model** — Splits a chosen portion of the goal equally across all territories. Useful for new product launches where historical data isn't predictive.

### 📊 Final Allocation
- Review the finalised territory goals in an interactive table and bar chart
- Reconciliation check to confirm the allocation matches the National Goal exactly
- Download the goal sheet as a professionally formatted Excel file

### 🔁 Back Testing
- Replay the Weighted Model against a past quarter to test its accuracy
- Compares simulated goals against actually-set goals and scores both against real sales
- Reports MAPE, RMSE, R², and a simulated payout comparison
- Includes a plain-language Diagnostic Readout that interprets the results and suggests next steps

---

## Required Input Format

The uploaded file must include the following columns:

| Column | Type | Description |
|---|---|---|
| `Week` | Date | Week start date |
| `Territory ID` | Text or Number | Unique territory identifier |
| `Territory Name` | Text | Territory display name |
| `Product` | Text | Product or product family name |
| `Sales` | Number | Actual sales for the week |
| `Goals` | Number | Goal set for the week |

---

## Running Locally

### 1. Clone or download this repository

```bash
git clone https://github.com/your-username/goal-setting-tool.git
cd goal-setting-tool
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Launch the app

```bash
streamlit run app.py
```

The app will open in your default browser at `http://localhost:8501`.

---

## Tech Stack

| Component | Used For |
|---|---|
| **Streamlit** | UI framework |
| **Pandas / NumPy** | Data processing and computation |
| **Plotly** | Interactive charts |
| **OpenPyXL** | Excel export with custom styling |
| **streamlit-extras** | Enhanced metric card styling |

---

## Workflow

```
1. Upload Data        →   Validate schema, clean missing/duplicate/outlier rows
2. Review Trends      →   Sales vs goals trend at weekly / monthly / quarterly level
3. Set National Goal  →   Enter the total goal for the upcoming period
4. Choose Model       →   Weighted · Fair Share · Equal Allocation
5. Review Allocation  →   Per-territory breakdown with growth targets
6. Back Test          →   Validate the model against a past quarter
7. Download           →   Export final goals as a formatted Excel file
```

---

## Notes

- The app defaults to **dark mode** (toggleable from the top-right of the navbar).
- Best results require **at least 6 months of weekly sales data**.
- The **Back Testing** module requires **at least 3 quarters** of historical data.
- All amounts are displayed in `$` by default. Currency symbol can be adjusted in the `fmt()` helper inside `app.py`.

---
