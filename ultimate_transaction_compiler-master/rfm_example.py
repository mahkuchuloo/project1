import pandas as pd
from datetime import datetime, timedelta

# Create sample data
data = {
    'Customer': ['A', 'B', 'C', 'D'],
    'Last_Purchase': [
        datetime.now() - timedelta(days=10),  # Very recent purchase
        datetime.now() - timedelta(days=100),  # Less recent purchase
        datetime.now() - timedelta(days=50),   # Moderately recent purchase
        datetime.now() - timedelta(days=200)   # Oldest purchase
    ],
    'Purchase_Count': [20, 5, 10, 2],         # Number of purchases
    'Total_Spent': [1000, 200, 500, 100]      # Total amount spent
}

df = pd.DataFrame(data)

# Calculate recency in days
df['Recency_Days'] = (datetime.now() - df['Last_Purchase']).dt.days

print("\n=== Customer Profiles ===")
print("\nCustomer A:")
print("- Most recent purchase (10 days ago)")
print("- Highest purchase frequency (20 purchases)")
print("- Highest total spend ($1,000)")
print("→ This should be our best customer")

print("\nCustomer B:")
print("- Less recent purchase (100 days ago)")
print("- Low purchase frequency (5 purchases)")
print("- Low total spend ($200)")
print("→ This is a declining customer")

print("\nCustomer C:")
print("- Moderately recent purchase (50 days ago)")
print("- Medium purchase frequency (10 purchases)")
print("- Medium total spend ($500)")
print("→ This is an average customer")

print("\nCustomer D:")
print("- Oldest purchase (200 days ago)")
print("- Lowest purchase frequency (2 purchases)")
print("- Lowest total spend ($100)")
print("→ This is at risk of churning")

print("\n=== Raw Data ===")
print(df)

print("\n=== Old Scoring Method (Incorrect) ===")
# Old method - all ascending=False
recency_old = df['Recency_Days'].rank(method='min', ascending=False).apply(lambda x: (x / len(df) * 10)).round(2)
frequency_old = df['Purchase_Count'].rank(method='min', ascending=False).apply(lambda x: (x / len(df) * 10)).round(2)
monetary_old = df['Total_Spent'].rank(method='min', ascending=False).apply(lambda x: (x / len(df) * 10)).round(2)

print("\nRecency Scores (Old):", recency_old.to_dict())
print("PROBLEM: Customer A (most recent) gets worst score (10.0) while Customer D (oldest) gets best score (2.5)")
print("         This is backwards - recent purchases should get better scores!")

print("\nFrequency Scores (Old):", frequency_old.to_dict())
print("This was correct - Customer A (20 purchases) gets best score (2.5) while Customer D (2 purchases) gets worst (10.0)")

print("\nMonetary Scores (Old):", monetary_old.to_dict())
print("This was correct - Customer A ($1,000) gets best score (2.5) while Customer D ($100) gets worst (10.0)")

print("\nTotal RFM Scores (Old):", (recency_old + frequency_old + monetary_old).to_dict())
print("PROBLEM: Customer D (worst customer) gets better total score (22.5) than Customer A (best customer, 15.0)")
print("         This is wrong because the recency scoring was reversed!")

print("\n=== New Scoring Method (Correct) ===")
# New method - recency ascending=True, others ascending=False
recency_new = df['Recency_Days'].rank(method='min',   .apply(lambda x: (x / len(df) * 10)).round(2)
frequency_new = df['Purchase_Count'].rank(method='min', ascending=False).apply(lambda x: (x / len(df) * 10)).round(2)
monetary_new = df['Total_Spent'].rank(method='min', ascending=False).apply(lambda x: (x / len(df) * 10)).round(2)

print("\nRecency Scores (New):", recency_new.to_dict())
print("CORRECT: Customer A (most recent) gets best score (2.5) while Customer D (oldest) gets worst score (10.0)")
print("         This properly rewards recent engagement!")

print("\nFrequency Scores (New):", frequency_new.to_dict())
print("CORRECT: Customer A (20 purchases) gets best score (2.5) while Customer D (2 purchases) gets worst (10.0)")
print("         This properly rewards higher purchase frequency!")

print("\nMonetary Scores (New):", monetary_new.to_dict())
print("CORRECT: Customer A ($1,000) gets best score (2.5) while Customer D ($100) gets worst (10.0)")
print("         This properly rewards higher spending!")

print("\nTotal RFM Scores (New):", (recency_new + frequency_new + monetary_new).to_dict())
print("\nWhy the new method is correct:")
print("1. Customer A (best customer):")
print("   - Gets best total score (7.5) because they are:")
print("   - Most recent (2.5 points)")
print("   - Highest frequency (2.5 points)")
print("   - Highest spending (2.5 points)")
print("   → This properly identifies them as the most valuable customer")

print("\n2. Customer D (worst customer):")
print("   - Gets worst total score (30.0) because they are:")
print("   - Least recent (10.0 points)")
print("   - Lowest frequency (10.0 points)")
print("   - Lowest spending (10.0 points)")
print("   → This properly identifies them as the least engaged customer")

print("\nThe new scoring method aligns with RFM best practices:")
print("1. Recency: Lower days since purchase = Better score (ascending=True)")
print("2. Frequency: Higher purchase count = Better score (ascending=False)")
print("3. Monetary: Higher total spent = Better score (ascending=False)")
