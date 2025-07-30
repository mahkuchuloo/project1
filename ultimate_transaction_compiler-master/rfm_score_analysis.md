# RFM Score Analysis

RFM (Recency, Frequency, Monetary) analysis is a customer segmentation technique that uses past purchase behavior to identify groups of customers and target them with tailored marketing strategies.

## Components

### 1. Recency Score (R)
- Measures how recently a customer has made a donation/transaction
- Based on the most recent transaction date
- Higher score (closer to 10) = more recent activity
- Lower score (closer to 0) = less recent activity
- **Why it matters**: Recently active donors are more likely to donate again compared to those who haven't donated in a long time

### 2. Frequency Score (F)
- Measures how often a customer makes donations/transactions
- Based on the count of unique transactions
- Higher score (closer to 10) = more frequent donations
- Lower score (closer to 0) = fewer donations
- **Why it matters**: Frequent donors show higher engagement and loyalty to the organization

### 3. Monetary Score (M)
- Measures how much money a customer has donated in total
- Based on the sum of all transaction amounts
- Higher score (closer to 10) = larger total donation amount
- Lower score (closer to 0) = smaller total donation amount
- **Why it matters**: Identifies high-value donors who contribute significantly to the organization's funding

## Combined RFM Score

The total RFM Score is calculated by adding the three component scores:
```
RFM Score = Recency Score + Frequency Score + Monetary Score
```

Range: 0-30 (each component contributes 0-10)

### Score Interpretation

#### High RFM Score (20-30)
- Best customers/donors
- Recently active
- Donate frequently
- High total contributions
- Strategy: Maintain relationship, VIP treatment

#### Medium RFM Score (10-20)
- Good customers/donors
- Moderately active
- Average donation frequency
- Moderate total contributions
- Strategy: Increase engagement, encourage more frequent donations

#### Low RFM Score (0-10)
- At-risk or lost customers/donors
- Haven't donated recently
- Infrequent donations
- Low total contributions
- Strategy: Re-engagement campaigns, special offers

## Scoring Method

Each component score is calculated using a ranking method:
1. Rank all customers for each metric
2. Convert rank to percentile (rank/count)
3. Scale to 0-10 range
4. Round to 2 decimal places

This ensures:
- Fair comparison across different scales
- Relative positioning of customers
- Easy interpretation on a 0-10 scale

## Using RFM Scores

RFM scores help in:
1. Customer Segmentation
2. Targeted Marketing
3. Resource Allocation
4. Campaign Planning
5. Donor Retention Strategies
6. Identifying High-Value Donors
7. Risk Assessment
8. Growth Opportunities

By understanding where each donor stands in terms of Recency, Frequency, and Monetary value, organizations can develop more effective, personalized engagement strategies.
