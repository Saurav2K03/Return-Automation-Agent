#!/usr/bin/env python3
"""
Faym Assignment — Data Analysis Solutions

Answers all 5 questions from the assignment:
  1. 7th Highest Debit Amount Through IMPS
  2. No. of transactions category-wise (UPI, IMPS, RTGS, NEFT)
  3. Bell Curve, Box Plot for load amount + stats
  4. Monthly Cohort view of active users (DEBIT txns)
  5. Top 10th percentile users by highest net amount (DEBIT − CREDIT)

Usage:
  source .venv/bin/activate
  python3 assignment_answers_clean.py
"""

import sqlite3
import os
import warnings

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

warnings.filterwarnings('ignore')

# Paths
DATA_FILE = os.path.join(os.path.dirname(__file__), 'Data Set (2).xlsx')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Load data
print("\nFAYM ASSIGNMENT — DATA ANALYSIS SOLUTIONS")

df = pd.read_excel(DATA_FILE, sheet_name='Table 1')

# Clean up column names (replace spaces with underscores for SQL)
df.columns = ['Transaction_Time', 'User_Id', 'Transaction_Amt', 'Narration',
              'Transaction_Type', 'Txn_id']
df['Transaction_Time'] = pd.to_datetime(df['Transaction_Time'])

print(f"\nDataset loaded: {len(df)} rows, {len(df.columns)} columns")
print(f"Date range: {df['Transaction_Time'].min().date()} to {df['Transaction_Time'].max().date()}")
print(f"Users: {sorted(df['User_Id'].unique())}")

# Load into SQLite so we can run SQL queries
conn = sqlite3.connect(':memory:')
df.to_sql('transactions', conn, index=False, if_exists='replace')

report_lines = []
report_lines.append("# Faym Assignment — Data Analysis Report\n")
report_lines.append("---\n")


# ---- Q1: 7th Highest Debit Amount Through IMPS ----

print("\nQ1: SQL for 7th Highest Debit Amount Through IMPS")

# Simple approach with LIMIT/OFFSET
q1_sql = """
SELECT Transaction_Amt, User_Id, Transaction_Time, Txn_id
FROM transactions
WHERE Narration = 'IMPS' AND Transaction_Type = 'DEBIT'
ORDER BY Transaction_Amt DESC
LIMIT 1 OFFSET 6;
"""

# Better approach with DENSE_RANK to handle ties properly
q1_sql_dense_rank = """
WITH RankedDebits AS (
    SELECT 
        Transaction_Amt,
        User_Id,
        Transaction_Time,
        Txn_id,
        DENSE_RANK() OVER (ORDER BY Transaction_Amt DESC) AS rnk
    FROM transactions
    WHERE Narration = 'IMPS' AND Transaction_Type = 'DEBIT'
)
SELECT Transaction_Amt, User_Id, Transaction_Time, Txn_id, rnk
FROM RankedDebits
WHERE rnk = 7;
"""

q1_result_simple = pd.read_sql_query(q1_sql, conn)
q1_result_dense = pd.read_sql_query(q1_sql_dense_rank, conn)

# Top 10 for context
q1_top10_sql = """
SELECT 
    DENSE_RANK() OVER (ORDER BY Transaction_Amt DESC) AS Rank,
    Transaction_Amt, User_Id, Transaction_Time, Txn_id
FROM transactions
WHERE Narration = 'IMPS' AND Transaction_Type = 'DEBIT'
ORDER BY Transaction_Amt DESC
LIMIT 10;
"""
q1_top10 = pd.read_sql_query(q1_top10_sql, conn)

print("\nSQL Query (Simple - using LIMIT OFFSET):")
print(q1_sql)
print("\nSQL Query (Using DENSE_RANK for tie handling):")
print(q1_sql_dense_rank)
print("\nTop 10 IMPS DEBIT amounts:")
print(q1_top10.to_string(index=False))
print(f"\n>>> 7th Highest IMPS DEBIT Amount: {q1_result_dense['Transaction_Amt'].values[0]}")

report_lines.append("## Question 1: 7th Highest Debit Amount Through IMPS\n")
report_lines.append("### SQL Query (Simple — LIMIT OFFSET)\n")
report_lines.append(f"```sql{q1_sql}```\n")
report_lines.append("### SQL Query (DENSE_RANK — handles ties)\n")
report_lines.append(f"```sql{q1_sql_dense_rank}```\n")
report_lines.append("### Result\n")
report_lines.append("**Top 10 IMPS DEBIT Amounts for context:**\n")
report_lines.append(q1_top10.to_markdown(index=False) + "\n\n")
report_lines.append(f"**Answer: The 7th highest DEBIT amount through IMPS is ₹{q1_result_dense['Transaction_Amt'].values[0]}**\n\n")
report_lines.append("---\n")


# ---- Q2: No. of transactions category-wise ----

print("\nQ2: SQL for No. of transactions category-wise (UPI, IMPS, RTGS, NEFT)")

q2_sql = """
SELECT 
    Narration AS Transaction_Category,
    COUNT(*) AS Total_Transactions,
    SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN 1 ELSE 0 END) AS Debit_Count,
    SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN 1 ELSE 0 END) AS Credit_Count
FROM transactions
WHERE Narration IN ('UPI', 'IMPS', 'RTGS', 'NEFT')
GROUP BY Narration
ORDER BY Total_Transactions DESC;
"""

q2_result = pd.read_sql_query(q2_sql, conn)

print("\nSQL Query:")
print(q2_sql)
print("\nResult:")
print(q2_result.to_string(index=False))

# Bonus: all categories including IFT with averages
q2_all_sql = """
SELECT 
    Narration AS Transaction_Category,
    COUNT(*) AS Total_Transactions,
    SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN 1 ELSE 0 END) AS Debit_Count,
    SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN 1 ELSE 0 END) AS Credit_Count,
    ROUND(AVG(Transaction_Amt), 2) AS Avg_Amount,
    SUM(Transaction_Amt) AS Total_Amount
FROM transactions
GROUP BY Narration
ORDER BY Total_Transactions DESC;
"""

q2_all_result = pd.read_sql_query(q2_all_sql, conn)
print("\nAll categories (including IFT):")
print(q2_all_result.to_string(index=False))

report_lines.append("## Question 2: Number of Transactions Category-wise\n")
report_lines.append("### SQL Query\n")
report_lines.append(f"```sql{q2_sql}```\n")
report_lines.append("### Result (UPI, IMPS, RTGS, NEFT)\n")
report_lines.append(q2_result.to_markdown(index=False) + "\n\n")
report_lines.append("### All Categories (including IFT)\n")
report_lines.append(f"```sql{q2_all_sql}```\n")
report_lines.append(q2_all_result.to_markdown(index=False) + "\n\n")
report_lines.append("---\n")


# ---- Q3: Bell Curve, Box Plot for load amount + stats ----

print("\nQ3: Bell Curve, Box Plot for load amount + Statistical Summary")

# "Load amount" = CREDIT transactions
credit_amounts = df[df['Transaction_Type'] == 'CREDIT']['Transaction_Amt']

stats = credit_amounts.describe()
stats_extended = pd.DataFrame({
    'Statistic': ['Count', 'Mean', 'Std Dev', 'Min', '25th Percentile (Q1)',
                  'Median (50th Percentile)', '75th Percentile (Q3)', 'Max',
                  'Skewness', 'Kurtosis', 'IQR', 'Variance'],
    'Value': [
        len(credit_amounts),
        round(credit_amounts.mean(), 2),
        round(credit_amounts.std(), 2),
        credit_amounts.min(),
        round(credit_amounts.quantile(0.25), 2),
        round(credit_amounts.median(), 2),
        round(credit_amounts.quantile(0.75), 2),
        credit_amounts.max(),
        round(credit_amounts.skew(), 4),
        round(credit_amounts.kurtosis(), 4),
        round(credit_amounts.quantile(0.75) - credit_amounts.quantile(0.25), 2),
        round(credit_amounts.var(), 2)
    ]
})

print("\nStatistical Summary (Credit/Load Amounts):")
print(stats_extended.to_string(index=False))

# Bell curve + Box plot side by side
fig, axes = plt.subplots(1, 2, figsize=(16, 7))

# Left: Histogram + normal distribution fit
ax1 = axes[0]
n, bins, patches = ax1.hist(credit_amounts, bins=25, density=True,
                            alpha=0.7, color='#4A90D9', edgecolor='white', linewidth=0.8,
                            label='Actual Distribution')

mu, sigma = credit_amounts.mean(), credit_amounts.std()
x = np.linspace(credit_amounts.min() - 500, credit_amounts.max() + 500, 200)
bell_curve = (1 / (sigma * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x - mu) / sigma) ** 2)

ax1.plot(x, bell_curve, color='#E74C3C', linewidth=2.5, label=f'Normal Fit (μ={mu:.0f}, σ={sigma:.0f})')
ax1.axvline(mu, color='#2ECC71', linestyle='--', linewidth=2, label=f'Mean = {mu:.0f}')
ax1.axvline(mu - sigma, color='#F39C12', linestyle=':', linewidth=1.5, alpha=0.7, label=f'μ ± σ')
ax1.axvline(mu + sigma, color='#F39C12', linestyle=':', linewidth=1.5, alpha=0.7)

ax1.set_title('Bell Curve — Distribution of Credit (Load) Amounts', fontsize=14, fontweight='bold', pad=15)
ax1.set_xlabel('Transaction Amount (₹)', fontsize=12)
ax1.set_ylabel('Density', fontsize=12)
ax1.legend(fontsize=10, loc='upper right')
ax1.grid(True, alpha=0.3)
ax1.set_facecolor('#FAFAFA')

# Right: Box plot
ax2 = axes[1]
bp = ax2.boxplot(credit_amounts, vert=True, widths=0.5, patch_artist=True,
                 boxprops=dict(facecolor='#4A90D9', alpha=0.7, linewidth=1.5),
                 medianprops=dict(color='#E74C3C', linewidth=2.5),
                 whiskerprops=dict(linewidth=1.5, color='#2C3E50'),
                 capprops=dict(linewidth=1.5, color='#2C3E50'),
                 flierprops=dict(marker='o', markerfacecolor='#E74C3C', markersize=6, alpha=0.6))

q1_val = credit_amounts.quantile(0.25)
q3_val = credit_amounts.quantile(0.75)
median_val = credit_amounts.median()

ax2.annotate(f'Q3 = {q3_val:.0f}', xy=(1.15, q3_val), fontsize=11, color='#2C3E50',
             fontweight='bold')
ax2.annotate(f'Median = {median_val:.0f}', xy=(1.15, median_val), fontsize=11, color='#E74C3C',
             fontweight='bold')
ax2.annotate(f'Q1 = {q1_val:.0f}', xy=(1.15, q1_val), fontsize=11, color='#2C3E50',
             fontweight='bold')
ax2.annotate(f'Min = {credit_amounts.min():.0f}', xy=(1.15, credit_amounts.min()), fontsize=10,
             color='#7F8C8D')
ax2.annotate(f'Max = {credit_amounts.max():.0f}', xy=(1.15, credit_amounts.max()), fontsize=10,
             color='#7F8C8D')

ax2.set_title('Box Plot — Credit (Load) Amounts', fontsize=14, fontweight='bold', pad=15)
ax2.set_ylabel('Transaction Amount (₹)', fontsize=12)
ax2.set_xticklabels(['Credit Amounts'])
ax2.grid(True, alpha=0.3, axis='y')
ax2.set_facecolor('#FAFAFA')

plt.tight_layout(pad=3.0)
plt.savefig(os.path.join(OUTPUT_DIR, 'q3_bell_curve_boxplot.png'), dpi=150, bbox_inches='tight',
            facecolor='white')
plt.close()
print(f"\n✅ Saved: {os.path.join(OUTPUT_DIR, 'q3_bell_curve_boxplot.png')}")

report_lines.append("## Question 3: Bell Curve & Box Plot for Load Amount\n")
report_lines.append("### Statistical Summary\n")
report_lines.append(stats_extended.to_markdown(index=False) + "\n\n")
report_lines.append("### Visualization\n")
report_lines.append(f"![Bell Curve and Box Plot](output/q3_bell_curve_boxplot.png)\n\n")
report_lines.append("### Interpretation\n")
report_lines.append(f"- The distribution of credit (load) amounts is approximately **uniform** with a slight negative skew ({credit_amounts.skew():.4f}).\n")
report_lines.append(f"- The kurtosis is {credit_amounts.kurtosis():.4f} (platykurtic — flatter than normal), consistent with the near-uniform spread.\n")
report_lines.append(f"- The IQR is ₹{q3_val - q1_val:.0f}, spanning from Q1=₹{q1_val:.0f} to Q3=₹{q3_val:.0f}.\n")
report_lines.append(f"- Mean (₹{mu:.0f}) and Median (₹{median_val:.0f}) are close, indicating a fairly symmetric distribution.\n\n")
report_lines.append("---\n")


# ---- Q4: Monthly Cohort View of Active Users (DEBIT txns) ----

print("\nQ4: Monthly Cohort View — Active Users (DEBIT Transactions)")

debit_df = df[df['Transaction_Type'] == 'DEBIT'].copy()
debit_df['txn_month'] = debit_df['Transaction_Time'].dt.to_period('M')

# Find each user's first DEBIT month (= their cohort)
first_debit_month = debit_df.groupby('User_Id')['txn_month'].min().reset_index()
first_debit_month.columns = ['User_Id', 'cohort_month']

debit_df = debit_df.merge(first_debit_month, on='User_Id')

user_active_months = debit_df.groupby(['User_Id', 'txn_month']).size().reset_index(name='txn_count')
user_active_months = user_active_months.merge(first_debit_month, on='User_Id')

all_months = sorted(debit_df['txn_month'].unique())
cohort_months = sorted(first_debit_month['cohort_month'].unique())

cohort_matrix = pd.DataFrame(index=[str(m) for m in cohort_months],
                             columns=[str(m) for m in all_months])
cohort_matrix = cohort_matrix.fillna(0)

for _, row in user_active_months.iterrows():
    cohort_key = str(row['cohort_month'])
    active_key = str(row['txn_month'])
    if cohort_key in cohort_matrix.index and active_key in cohort_matrix.columns:
        cohort_matrix.loc[cohort_key, active_key] = int(cohort_matrix.loc[cohort_key, active_key]) + 1

# Build proper cohort matrix counting unique users (not txn counts)
cohort_data = []
for cm in cohort_months:
    cohort_users = first_debit_month[first_debit_month['cohort_month'] == cm]['User_Id'].tolist()
    row_data = {'First Month': str(cm)}
    for am in all_months:
        active_users = debit_df[(debit_df['User_Id'].isin(cohort_users)) & 
                                (debit_df['txn_month'] == am)]['User_Id'].nunique()
        row_data[str(am)] = active_users
    cohort_data.append(row_data)

cohort_df = pd.DataFrame(cohort_data)
cohort_df = cohort_df.set_index('First Month')

# Blank out cells where active month < cohort month (users can't be active before joining)
cohort_df = cohort_df.astype(object)
for i, cm in enumerate(cohort_months):
    for am in all_months:
        if am < cm:
            cohort_df.loc[str(cm), str(am)] = ''

print("\nMonthly Cohort View (Number of Active Users doing DEBIT txns):")
print(cohort_df.to_string())

print("\n\nCohort Sizes (users who made their first DEBIT in each month):")
cohort_sizes = first_debit_month.groupby('cohort_month').size()
print(cohort_sizes.to_string())

# Heatmap
fig, ax = plt.subplots(figsize=(12, 7))

cohort_numeric = cohort_df.copy()
cohort_numeric = cohort_numeric.replace('', np.nan).astype(float)

im = ax.imshow(cohort_numeric.values.astype(float), cmap='YlOrRd', aspect='auto',
               interpolation='nearest')

ax.set_xticks(range(len(all_months)))
ax.set_xticklabels([str(m) for m in all_months], rotation=45, ha='right', fontsize=11)
ax.set_yticks(range(len(cohort_months)))
ax.set_yticklabels([str(m) for m in cohort_months], fontsize=11)

# Annotate cells with values
for i in range(len(cohort_months)):
    for j in range(len(all_months)):
        val = cohort_numeric.iloc[i, j]
        if not np.isnan(val) and val > 0:
            text_color = 'white' if val > cohort_numeric.max().max() * 0.6 else 'black'
            ax.text(j, i, f'{int(val)}', ha='center', va='center',
                    fontsize=13, fontweight='bold', color=text_color)

ax.set_title('Monthly Cohort View — Active Users (DEBIT Transactions)',
             fontsize=14, fontweight='bold', pad=15)
ax.set_xlabel('Active Month', fontsize=12, labelpad=10)
ax.set_ylabel('First DEBIT Month (Cohort)', fontsize=12, labelpad=10)

cbar = plt.colorbar(im, ax=ax, shrink=0.8)
cbar.set_label('Number of Active Users', fontsize=11)

plt.tight_layout()
plt.savefig(os.path.join(OUTPUT_DIR, 'q4_cohort_heatmap.png'), dpi=150, bbox_inches='tight',
            facecolor='white')
plt.close()
print(f"\n✅ Saved: {os.path.join(OUTPUT_DIR, 'q4_cohort_heatmap.png')}")

report_lines.append("## Question 4: Monthly Cohort View — Active Users (DEBIT Txns)\n")
report_lines.append("A cohort is defined by the **first month** a user made a DEBIT transaction. ")
report_lines.append("Each cell shows how many users from that cohort were active (made DEBIT txns) in the given month.\n\n")
report_lines.append("### Cohort Matrix\n")
report_lines.append(cohort_df.reset_index().to_markdown(index=False) + "\n\n")
report_lines.append("### Cohort Sizes\n")
report_lines.append("| Cohort Month | Users |\n|---|---|\n")
for idx, val in cohort_sizes.items():
    report_lines.append(f"| {idx} | {val} |\n")
report_lines.append("\n")
report_lines.append("### Heatmap\n")
report_lines.append(f"![Cohort Heatmap](output/q4_cohort_heatmap.png)\n\n")
report_lines.append("---\n")


# ---- Q5: Top 10th Percentile Users by Highest Net Amount ----

print("\nQ5: SQL for Top 10th Percentile Users — Highest Net Amount (DEBIT − CREDIT)")

# Using NTILE(10) — divides users into 10 buckets, top bucket = top 10%
q5_sql = """
WITH UserNetAmount AS (
    SELECT 
        User_Id,
        SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) AS Total_Debit,
        SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Total_Credit,
        SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) -
        SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Net_Amount
    FROM transactions
    GROUP BY User_Id
),
PercentileCalc AS (
    SELECT 
        User_Id,
        Total_Debit,
        Total_Credit,
        Net_Amount,
        NTILE(10) OVER (ORDER BY Net_Amount DESC) AS decile
    FROM UserNetAmount
)
SELECT User_Id, Net_Amount, Total_Debit, Total_Credit, decile
FROM PercentileCalc
WHERE decile = 1
ORDER BY Net_Amount DESC;
"""

# Alternative with PERCENT_RANK (more precise for non-integer boundaries)
q5_sql_alt = """
WITH UserNetAmount AS (
    SELECT 
        User_Id,
        SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) AS Total_Debit,
        SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Total_Credit,
        SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) -
        SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Net_Amount
    FROM transactions
    GROUP BY User_Id
),
PercentileCalc AS (
    SELECT 
        User_Id,
        Total_Debit,
        Total_Credit,
        Net_Amount,
        PERCENT_RANK() OVER (ORDER BY Net_Amount ASC) AS pct_rank
    FROM UserNetAmount
)
SELECT User_Id, Net_Amount, Total_Debit, Total_Credit,
       ROUND(pct_rank * 100, 2) AS Percentile_Rank
FROM PercentileCalc
WHERE pct_rank >= 0.90
ORDER BY Net_Amount DESC;
"""

q5_result = pd.read_sql_query(q5_sql, conn)
q5_result_alt = pd.read_sql_query(q5_sql_alt, conn)

# All users for context
q5_all_sql = """
SELECT 
    User_Id,
    SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) AS Total_Debit,
    SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Total_Credit,
    SUM(CASE WHEN Transaction_Type = 'DEBIT' THEN Transaction_Amt ELSE 0 END) -
    SUM(CASE WHEN Transaction_Type = 'CREDIT' THEN Transaction_Amt ELSE 0 END) AS Net_Amount
FROM transactions
GROUP BY User_Id
ORDER BY Net_Amount DESC;
"""
q5_all = pd.read_sql_query(q5_all_sql, conn)

print("\nSQL Query (Using NTILE):")
print(q5_sql)
print("\nSQL Query (Using PERCENT_RANK):")
print(q5_sql_alt)
print("\nAll Users — Net Amount (DEBIT − CREDIT):")
print(q5_all.to_string(index=False))
print("\n>>> Top 10th Percentile Users:")
print(q5_result_alt.to_string(index=False))

report_lines.append("## Question 5: Top 10th Percentile Users — Highest Net Amount\n")
report_lines.append("Net Amount = Total DEBIT − Total CREDIT per user.\n\n")
report_lines.append("### SQL Query (Using NTILE)\n")
report_lines.append(f"```sql{q5_sql}```\n")
report_lines.append("### SQL Query (Using PERCENT_RANK)\n")
report_lines.append(f"```sql{q5_sql_alt}```\n")
report_lines.append("### All Users — Net Amount Ranking\n")
report_lines.append(q5_all.to_markdown(index=False) + "\n\n")
report_lines.append("### Top 10th Percentile Result\n")
report_lines.append(q5_result_alt.to_markdown(index=False) + "\n\n")
report_lines.append("---\n")

conn.close()

# Write report
report_path = os.path.join(os.path.dirname(__file__), 'assignment_report.md')
with open(report_path, 'w') as f:
    f.writelines(report_lines)

print(f"\n✅ Report saved to: {report_path}")
print(f"✅ Visualizations saved to: {OUTPUT_DIR}/")
