# Fetching and Analyzing Top 50 Live Cryptocurrency Data

This repository contains a Python script to fetch live cryptocurrency data for the top 50 cryptocurrencies by market capitalization, analyze the data, and generate a live-updating Excel sheet and a PDF analysis report.

---

## Files in the Repository

1. **`crypto_tracker.py`**  
   - A Python script that fetches live cryptocurrency data from the CoinGecko API.
   - Performs analysis on the data (e.g., top 5 cryptocurrencies by market cap, average price, highest/lowest 24h price change).
   - Saves the data to an Excel file.

2. **`crypto_data.xlsx`**  
   - An Excel file containing two sheets:
     - **Cryptocurrency Data**: Live data for the top 50 cryptocurrencies.
     - **Analysis Results**: Key insights and a separate table for the top 5 cryptocurrencies by market cap.

3. **`Cryptocurrency_Analysis_Report.pdf`**  
   - A PDF report summarizing the key insights and analysis from the fetched data.

---

## How to Use

### Prerequisites
- Python 3.x installed on your system.
- Required Python libraries: `requests`, `pandas`, `openpyxl`, `fpdf`.

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/your-repo-name.git

2. Navigate to the project directory:
   cd your-repo-name

3. Install the required Python libraries:
   pip install requests pandas openpyxl
   
