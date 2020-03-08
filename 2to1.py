# This example extracts data from 2 excel files and produces a new file

import pandas as pd
import matplotlib.pyplot as plt

text_data = {}
with pd.ExcelFile("data/movies-text.xlsx") as text_file:
    # Read third sheet (index starts from 0)
    text_data = pd.read_excel(text_file, sheet_name=2)
    # Remove rows with missing title or year. inplace=Trure: modify text_date directly.
    text_data.dropna(subset=['Title', 'Year'], inplace=True)
    # Remove white spaces around title text
    text_data['Title'] = text_data['Title'].str.strip()
    # Make the combination of Title and Year the index
    text_data.set_index(['Title', 'Year'], inplace=True)

number_data = {}
with pd.ExcelFile("data/movies-numbers.xlsx") as numbers_file:
    # Read sheet titled '2010s'. Sheets can be referred to by either their index (starting from 0)
    # or by their names, like '2010s'
    number_data = pd.read_excel(numbers_file, sheet_name='2010s')
    number_data.dropna(subset=['Title', 'Year'], inplace=True)
    number_data['Title'] = number_data['Title'].str.strip()
    number_data.set_index(['Title', 'Year'], inplace=True)

# Combine data from 2 sheets
combined_data = text_data.join(number_data)
# If needed, we can remove duplicate values in the data set
combined_data.drop_duplicates(True)

# Let's create a new column of data
combined_data["Net Earnings"] = combined_data["Gross Earnings"] - combined_data["Budget"]

# Now let's do some data analysis
# Sort the movies gross earnings
sorted_by_gross = combined_data.sort_values(by=['Gross Earnings'], ascending=False)
# Print the top 5 movies with highest earnings and their IMDB score
print(sorted_by_gross[['Gross Earnings', 'Net Earnings', 'IMDB Score']].head(5))

# What if I want to know, on a yearly basis, the average IMDB scores for all the movies in that year?
print(combined_data.pivot_table(index=['Year'])['IMDB Score'])

# Sort movies in ascending order, by year first, then by title.
combined_data.sort_values(['Year', 'Title'], inplace=True)

# Write data to merged file.
with pd.ExcelWriter('data/merged-data.xlsx') as writer:
    combined_data.to_excel(writer, sheet_name='2010s')
