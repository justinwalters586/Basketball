import pandas as pd

# Load the spreadsheet
df = pd.read_excel('C:/Users/WaltersJ07/Documents/March-Madness.xlsx')

# Define the weights for each statistic
weights = {
    'Adj. Off. Eff.': 0.15,
    'Adj. Def. Eff.': 0.15,
    'Free Throw %': 0.1,
    'Turnover Margin': 0.1,
    'Effective Field Goal Percentage': 0.1,
    'Three Point Percentage': 0.1,
    'Field Goal Percentage Defense': 0.1,
    'Rebounding Margin': 0.1,
    'Opponent Effective Possesion Ratio': 0.05,
    'Effective Possesion Ratio': 0.05,
    'Extra Scoring Chances per Game': 0.05
}

# Check if all statistic columns exist in the dataframe
for stat in weights.keys():
    if stat not in df.columns:
        print(f"The '{stat}' column does not exist in the spreadsheet. Please check the column name.")
        exit()

# Calculate the final score for each team
df['Score'] = sum(df[stat] * weight for stat, weight in weights.items())

# Rank the teams based on the final score
df['Rank'] = df['Score'].rank(ascending=False, method='min')

# Sort the dataframe by rank
df = df.sort_values('Rank')

# Print the ranked teams
print(df[['Team', 'Rank']])