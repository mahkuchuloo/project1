import pandas as pd
import numpy as np
from typing import Callable

class RFMScorer:
    """
    A class containing different methods for calculating RFM scores.
    Each method ensures scores are between 1-10.
    
    RFM (Recency, Frequency, Monetary) scoring is a customer segmentation technique that uses past purchase behavior 
    to segment customers. This class provides multiple methods to calculate these scores:
    
    1. Percentile Scoring (Original VBA Method):
       - Ranks each customer relative to all other customers
       - Converts ranks to percentiles and scales to 1-10
       - Advantage: Maintains relative positioning of customers
       - Best for: When you want to know how each customer ranks compared to others
       - Example: If a customer ranks 50th out of 100 customers, their score would be 5.0
    
    2. Quartile Scoring:
       - Divides customers into 4 equal groups
       - Assigns fixed scores: 1, 4, 7, or 10 to each group
       - Advantage: Creates clear segments with meaningful gaps between them
       - Best for: When you want distinct customer tiers
       - Example: Top 25% get score 10, next 25% get 7, next 25% get 4, bottom 25% get 1
    
    3. Equal-width Scoring:
       - Divides the value range (max-min) into 10 equal intervals
       - Assigns scores based on which interval the value falls into
       - Advantage: Scores directly reflect the actual values
       - Best for: When the actual values should determine the score
       - Example: If values range 0-100, intervals would be 0-10, 10-20, etc.
    
    4. Threshold Scoring:
       - Uses predefined business thresholds to assign scores
       - Scores based on crossing specific value thresholds
       - Advantage: Aligns with business-defined goals or expectations
       - Best for: When you have specific target values
       - Example: Score 10 for >$1000, 7 for >$500, 4 for >$100, 1 for rest
    
    5. Z-score Scoring:
       - Converts values to standard deviations from mean
       - Maps z-scores from -3 to +3 to range 1-10
       - Advantage: Accounts for statistical distribution of values
       - Best for: Normally distributed data
       - Example: Values 1 std dev above mean get scores around 7-8
    
    6. Logarithmic Scoring:
       - Applies log transformation before scoring
       - Compresses high values and expands low values
       - Advantage: Handles skewed distributions and outliers
       - Best for: When values have exponential distribution
       - Example: Difference between $10-$100 treated similar to $100-$1000
    """
    
    @staticmethod
    def preprocess_recency(series: pd.Series) -> pd.Series:
        """
        Preprocess recency dates by converting them to days between each purchase
        and the latest purchase date.
        
        Parameters:
        -----------
        series : pd.Series
            Series of datetime values representing purchase dates
            
        Returns:
        --------
        pd.Series of integers representing days between each purchase and the latest purchase
        """
        if not pd.api.types.is_datetime64_any_dtype(series):
            return series
            
        latest_date = series.max()
        days_diff = (latest_date - series).dt.days
        return days_diff
    
    @staticmethod
    def percentile_scoring(series: pd.Series, ascending: bool = True) -> pd.Series:
        """
        Original VBA method using RANK.EQ/COUNT*10 formula.
        This is the exact implementation from the VBA code:
        =IFERROR(RANK.EQ(value,column,1)/COUNT(column)*10,0)
        
        How it works:
        1. Ranks each value (RANK.EQ in VBA, rank() in pandas)
        2. Divides by count of values to get percentile
        3. Multiplies by 10 to get score in 1-10 range
        4. Uses method='min' to match VBA's RANK.EQ behavior
        
        Example:
        For values [100, 200, 300, 400, 500]:
        - Ranks would be [1, 2, 3, 4, 5]
        - Divided by count (5) = [0.2, 0.4, 0.6, 0.8, 1.0]
        - Times 10 = [2, 4, 6, 8, 10]
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        rank = series.rank(method='min', ascending=ascending)
        count = len(series)
        scores = ((rank / count) * 10).round(2)
        return scores.clip(1, 10)

    @staticmethod
    def quartile_scoring(series: pd.Series, ascending: bool = True) -> pd.Series:
        """
        Quartile-based scoring method.
        
        How it works:
        1. Uses pandas qcut to divide data into 4 equal-sized groups
        2. Maps each quartile to a fixed score:
           - Bottom quartile (0-25%): Score 1
           - Lower middle quartile (25-50%): Score 4
           - Upper middle quartile (50-75%): Score 7
           - Top quartile (75-100%): Score 10
        
        Example:
        For values [1,2,3,4,5,6,7,8]:
        - Q1 (1,2): Score 1
        - Q2 (3,4): Score 4
        - Q3 (5,6): Score 7
        - Q4 (7,8): Score 10
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        quartiles = pd.qcut(series, q=4, labels=False, duplicates='drop')
        
        if ascending:
            score_map = {0: 1, 1: 4, 2: 7, 3: 10}
        else:
            score_map = {0: 10, 1: 7, 2: 4, 3: 1}
            
        scores = quartiles.map(score_map)
        return scores.fillna(1)

    @staticmethod
    def equal_width_scoring(series: pd.Series, ascending: bool = True) -> pd.Series:
        """
        Equal-width binning method.
        
        How it works:
        1. Finds the range of values (max - min)
        2. Divides this range into 10 equal-width bins
        3. Assigns scores 1-10 based on which bin the value falls into
        
        Example:
        For values ranging 0-100:
        - Bin 1 (0-10): Score 1
        - Bin 2 (10-20): Score 2
        ...
        - Bin 10 (90-100): Score 10
        
        This method is sensitive to outliers as they can stretch the bins.
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        bins = pd.cut(series, bins=10, labels=False, duplicates='drop')
        
        if ascending:
            scores = bins + 1
        else:
            scores = 10 - bins
            
        return scores.fillna(1)

    @staticmethod
    def threshold_scoring(series: pd.Series, thresholds: list, ascending: bool = True) -> pd.Series:
        """
        Custom threshold-based scoring method.
        
        How it works:
        1. Takes list of threshold values defined by business rules
        2. Assigns scores based on where value falls in thresholds
        3. Score increases as thresholds are crossed
        
        Example with monetary thresholds [100, 500, 1000]:
        - Value <= 100: Score 1
        - Value <= 500: Score 2
        - Value <= 1000: Score 3
        - Value > 1000: Score 10
        
        This method lets business logic directly drive scoring.
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        scores = pd.Series(index=series.index, dtype=float)
        
        if ascending:
            thresholds = sorted(thresholds)
            for i, threshold in enumerate(thresholds, 1):
                scores[series <= threshold] = i
            scores[series > thresholds[-1]] = 10
        else:
            thresholds = sorted(thresholds, reverse=True)
            for i, threshold in enumerate(thresholds, 1):
                scores[series >= threshold] = i
            scores[series < thresholds[-1]] = 10
            
        return scores.fillna(1)

    @staticmethod
    def zscore_scoring(series: pd.Series, ascending: bool = True) -> pd.Series:
        """
        Z-score based scoring method.
        
        How it works:
        1. Calculates z-score: (value - mean) / std_dev
        2. Maps z-scores from -3 to +3 to range 1-10
        3. Clips outliers to ensure 1-10 range
        
        Example:
        - Mean value gets score ~5.5
        - 1 std dev above mean gets score ~7.2
        - 2 std dev above mean gets score ~8.8
        - 3 std dev above mean gets score 10
        
        Best for normally distributed data where mean/std_dev are meaningful.
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        z_scores = (series - series.mean()) / series.std()
        
        if ascending:
            scores = (z_scores + 3) * (10/6)
        else:
            scores = (-z_scores + 3) * (10/6)
            
        return scores.clip(1, 10)

    @staticmethod
    def logarithmic_scoring(series: pd.Series, ascending: bool = True) -> pd.Series:
        """
        Logarithmic scoring method.
        
        How it works:
        1. Applies log1p transformation: log(x + 1)
        2. Scales transformed values to 1-10 range
        3. Handles skewed distributions by compressing high values
        
        Example effect on differences:
        Original    Log1p       Score
        1 -> 10     0.7 -> 2.4   ~2 points
        10 -> 100   2.4 -> 4.6   ~3 points
        100 -> 1000 4.6 -> 6.9   ~2 points
        
        Best for exponentially distributed data like monetary values.
        The log transformation makes relative differences more important
        than absolute differences.
        """
        if len(series) == 0:
            return pd.Series([])
            
        # Preprocess recency dates if needed
        if pd.api.types.is_datetime64_any_dtype(series):
            series = RFMScorer.preprocess_recency(series)
        
        min_val = series.min()
        if min_val <= 0:
            adjusted_series = series - min_val + 1
        else:
            adjusted_series = series
            
        log_scores = np.log1p(adjusted_series)
        
        min_score = log_scores.min()
        max_score = log_scores.max()
        scores = 1 + 9 * (log_scores - min_score) / (max_score - min_score)
        
        if not ascending:
            scores = 11 - scores
            
        return scores.clip(1, 10)

def calculate_rfm_scores(df: pd.DataFrame, 
                        scoring_method: Callable[[pd.Series, bool], pd.Series] = RFMScorer.percentile_scoring
                        ) -> pd.DataFrame:
    """
    Calculate RFM scores using the specified scoring method.
    
    Parameters:
    -----------
    df : DataFrame
        Must contain 'Recency Criteria', 'Frequency Criteria', and 'Monetary Criteria' columns
    scoring_method : Callable
        The scoring method to use from RFMScorer class
        Default is percentile_scoring which matches the original VBA implementation
        
    Returns:
    --------
    DataFrame with added RFM score columns
    
    Example:
    --------
    # Basic usage with default (percentile) scoring
    rfm_scores = calculate_rfm_scores(customer_data)
    
    # Using a different scoring method
    rfm_scores = calculate_rfm_scores(customer_data, RFMScorer.quartile_scoring)
    
    # Using threshold scoring with custom thresholds
    def custom_threshold_scoring(series, ascending):
        thresholds = [100, 500, 1000, 5000, 10000]
        return RFMScorer.threshold_scoring(series, thresholds, ascending)
    
    rfm_scores = calculate_rfm_scores(customer_data, custom_threshold_scoring)
    """
    # Create copy to avoid modifying original
    result = df.copy()
    
    # Calculate individual RFM scores
    # Note: Recency is reversed (more recent = higher score)
    result['Recency Score'] = scoring_method(result['Recency Criteria'], ascending=False)
    result['Frequency Score'] = scoring_method(result['Frequency Criteria'], ascending=True)
    result['Monetary Score'] = scoring_method(result['Monetary Criteria'], ascending=True)
    
    # Calculate total RFM score
    result['RFM Score'] = (result['Recency Score'] + 
                          result['Frequency Score'] + 
                          result['Monetary Score'])
    
    return result
