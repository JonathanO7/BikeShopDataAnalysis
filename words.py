
# Import pandas library and read Excel file
import pandas as pd

excel_file = pd.ExcelFile('VI_New_raw_data_update_final.xlsx')

# Create an empty dictionary to store column values for each sheet
sheet_column_values = {}

# Loop over all sheets in the Excel file
for sheet_name in excel_file.sheet_names:
    # Read the sheet into a DataFrame, skipping the first row
    df = pd.read_excel(excel_file, sheet_name=sheet_name, skiprows=1)
    # Create an empty dictionary to store column values for this sheet
    column_values = {}
    # Loop over the columns in the DataFrame
    for column in df.columns:
        # Store the values in the column as a list
        values = df[column].tolist()
        # Add the list to the dictionary with the column name as the key
        column_values[column] = values
    # Add the column values dictionary to the sheet dictionary with the sheet name as the key
    sheet_column_values[sheet_name] = column_values



import matplotlib.pyplot as plt
from wordcloud import WordCloud


CustomerDemographic = sheet_column_values["CustomerDemographic"]


job_title = CustomerDemographic["job_title"]

def filter_valid_strings(elements):
    valid_strings = [str(element).strip() for element in elements if isinstance(element, str) and str(element).strip()]
    return valid_strings



def remove_substrings(job_title):

    filtered_list = []
    for string in job_title:
        filtered_string = string.replace("II", "").replace("III", "").replace("I", "").replace("IV", "").replace("V", "")
        filtered_list.append(filtered_string)
    return filtered_list


def generate_word_cloud(job_title):
    job_title = filter_valid_strings(job_title)
    job_title = remove_substrings(job_title)
    

    # Join all the names into a single string
    text = ' '.join(job_title)

    # Create a WordCloud object
    wordcloud = WordCloud(width=800, height=400, background_color='white').generate(text)

    # Display the word cloud using matplotlib
    plt.figure(figsize=(10, 5))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')
    plt.show()

# Example usage

generate_word_cloud(job_title)

