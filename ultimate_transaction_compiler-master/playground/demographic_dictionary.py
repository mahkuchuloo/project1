import pandas as pd
import numpy as np
import random

# The list of IDs
ids = [
    109185848, 109188607, 108849492, 109177804, 108966408, 108851191, 109185249,
    109173842, 109164167, 109163010, 109194838, 109162950, 108966409, 108840440,
    109173843, 108857644, 109194839, 109151060, 109154190, 109178629, 108848243,
    108840441, 109183584, 108966410, 109186216, 109157227, 109160369, 109173844,
    108842916, 109062178, 109239919, 109222288, 109377802, 109210754, 109163116,
    109162550, 109169894, 109214722, 109214677, 108859591, 109275835, 109214737,
    109189135, 109194828, 109161059, 108870756, 109239803, 109302903, 109244742,
    109219833, 109244971, 108875489, 109246395, 109220418, 109239974, 109239940,
    109307880, 109239819, 109307882, 109307881, 109173846, 109307883, 109239268,
    109307884, 109307885, 109307886, 109307889, 109307894, 109307895, 109307899,
    109307892, 109307897, 109307898, 109307443, 109307900, 109307902, 109307160,
    109307887, 109307029, 108876836, 109307890, 109307891, 109307893, 109307896,
    109307484, 109307901, 109307888, "maaj101@houmaic.com", 109191571, 109307903,
    109307905, 109307904, 109307907, 109307906, 109307256, 109306932, 109307908,
    109307911, 109307909
]

# Create random but realistic data
n_rows = len(ids)
ages = [random.randint(43, 89) for _ in range(n_rows)]
genders = random.choices(['male', 'female'], k=n_rows)
races = random.choices(['caucasian', 'black'], weights=[0.8, 0.2], k=n_rows)

# Function to determine age range
def get_age_range(age):
    if age >= 78:
        return 'Silent Generation: Age 78 and above'
    elif 59 <= age <= 77:
        return 'Baby Boomer: Age 59 to 77'
    elif 47 <= age <= 58:
        return 'Generation X: Age 47 to 58'
    elif 43 <= age <= 46:
        return 'Xennial: Age 43 to 46'
    else:
        return 'None'

# Create the DataFrame
data = {
    'ID': ids,
    'AGE': ages,
    'GENDER': genders,
    'RACE': races
}

df = pd.DataFrame(data)
df['AGE RANGE'] = df['AGE'].apply(get_age_range)

# Save to Excel
df.to_excel('playground/Demographic dictionary.xlsx', index=False)
print("Excel file 'Demographic dictionary.xlsx' has been created successfully!")
