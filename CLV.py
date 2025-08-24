import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# -- Step 1: Load Data --
# Load the Online Retail dataset
try:
    df = pd.read_excel(r'c:\Users\rudra\Downloads\practise\Online Retail.xlsx')
except FileNotFoundError:
    print("Error: 'Online Retail.xlsx' not found. Please make sure the file is in the 'practise' directory.")
    exit()

# -- Step 2: Data Cleaning and Preparation --
# Remove rows with missing CustomerID
df.dropna(subset=['CustomerID'], inplace=True)
# Convert CustomerID to integer
df['CustomerID'] = df['CustomerID'].astype(int)
# Remove returns (InvoiceNo starts with 'C')
df = df[~df['InvoiceNo'].astype(str).str.startswith('C')]
# Calculate SaleAmount
df['SaleAmount'] = df['Quantity'] * df['UnitPrice']

# -- Step 3: Calculate RFM Values --
# Set a fixed date for analysis (last date in the dataset + 1 day)
today = df['InvoiceDate'].max() + pd.Timedelta(days=1)

rfm = df.groupby('CustomerID').agg({
    'InvoiceDate': lambda date: (today - date.max()).days,  # Recency
    'InvoiceNo': 'nunique',                               # Frequency
    'SaleAmount': 'sum'                                   # Monetary
}).rename(columns={
    'InvoiceDate': 'Recency',
    'InvoiceNo': 'Frequency',
    'SaleAmount': 'Monetary'
})

# Filter out customers with zero or negative monetary value (often returns or errors)
rfm = rfm[rfm['Monetary'] > 0]

# -- Step 4: Create RFM Scores --
# Use quintiles to score each metric from 1 to 5
rfm['R_Score'] = pd.qcut(rfm['Recency'], 5, labels=[5, 4, 3, 2, 1])
rfm['F_Score'] = pd.qcut(rfm['Frequency'].rank(method='first'), 5, labels=[1, 2, 3, 4, 5])
rfm['M_Score'] = pd.qcut(rfm['Monetary'], 5, labels=[1, 2, 3, 4, 5])

# -- Step 5: Create Final RFM Segment --
# Combine the scores
rfm['RFM_Score'] = rfm['R_Score'].astype(str) + rfm['F_Score'].astype(str) + rfm['M_Score'].astype(str)

# Define segmentation rules
def segment_customer(row):
    if row['R_Score'] == 5 and row['F_Score'] >= 4 and row['M_Score'] >= 4:
        return 'Champions'
    elif row['R_Score'] >= 4 and row['F_Score'] >= 4:
        return 'Loyal Customers'
    elif row['R_Score'] >= 3 and row['F_Score'] >= 3 and row['M_Score'] >= 3:
        return 'Potential Loyalists'
    elif row['R_Score'] <= 2 and row['F_Score'] <= 2:
        return 'At-Risk'
    elif row['R_Score'] <= 2:
        return 'Needs Attention'
    else:
        return 'Standard'

rfm['Segment'] = rfm.apply(segment_customer, axis=1)

# -- Step 6: Calculate Average CLV (Monetary) per Segment --
avg_clv_per_segment = rfm.groupby('Segment')['Monetary'].mean().sort_values(ascending=False)

# -- Step 7: Generate and Save Visualizations --
# RFM Segment Distribution
plt.figure(figsize=(12, 6))
sns.countplot(x='Segment', data=rfm, order=rfm['Segment'].value_counts().index)
plt.title('Customer Segmentation Distribution (RFM)')
plt.xlabel('Segment')
plt.ylabel('Number of Customers')
plt.xticks(rotation=45)
plt.tight_layout()
rfm_visualization_path = r'c:\Users\rudra\Downloads\practise\CLV ANALYSIS\customer_segmentation_rfm.png'
plt.savefig(rfm_visualization_path)
print(f"RFM segmentation chart saved to {rfm_visualization_path}")

# -- Step 8: Generate and Save Report --
output_path = r'c:\Users\rudra\Downloads\practise\CLV ANALYSIS\clv_analysis_report.txt'
with open(output_path, 'w') as f:
    f.write("Customer Segmentation Analysis Report\n")
    f.write("="*40 + "\n\n")

    f.write("1. RFM Segmentation Summary:\n\n")
    rfm_output = rfm[['Recency', 'Frequency', 'Monetary', 'Segment']].rename(columns={
        'Recency': 'Recency (Days)',
        'Frequency': 'Frequency (No. of Orders)',
        'Monetary': 'Monetary ($)'
    })
    f.write(rfm_output.to_string())
    f.write("\n\n" + "="*40 + "\n\n")

    f.write("2. Average Monetary Value (CLV Proxy) per RFM Segment:\n\n")
    f.write(avg_clv_per_segment.to_string())
    f.write("\n\n" + "="*40 + "\n\n")

    f.write("3. RFM Segment Distribution:\n\n")
    f.write(rfm['Segment'].value_counts().to_string())
    f.write("\n\n" + "="*40 + "\n\n")

    f.write("4. Actionable Insights and Recommendations (RFM-based):\n\n")
    insights = {
        "Champions": "These are your best and most loyal customers. Reward them with exclusive offers, early access to new products, and loyalty programs. They can also be your brand ambassadors.",
        "Loyal Customers": "These customers buy frequently. Nurture them to become Champions. Offer them loyalty points and involve them in customer surveys to make them feel valued.",
        "Potential Loyalists": "These are recent customers with average frequency and spending. Engage them with personalized marketing campaigns and offer them membership or loyalty programs.",
        "At-Risk": "These customers have not purchased in a long time and have low frequency. Reach out to them with personalized reactivation campaigns and special discounts to win them back.",
        "Needs Attention": "These customers have low recency, but also low frequency and monetary value. Try to understand their needs better through surveys and offer them relevant products.",
        "Standard": "This is a mixed group. While some may not be very profitable, it's good to keep them engaged with generic marketing campaigns and newsletters."
    }
    for segment, insight in insights.items():
        f.write(f"- {segment}:\n{insight}\n\n")

print(f"CLV segmentation report saved to {output_path}")