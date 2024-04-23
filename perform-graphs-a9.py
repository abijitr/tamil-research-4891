import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def clean_numeric_columns(data, columns_to_clean):
    # Iterate through each column
    for col in columns_to_clean:
        # Convert column to numeric, coerce non-numeric values to NaN
        data[col] = pd.to_numeric(data[col], errors='coerce')

def rename_columns(data):
    # Rename columns containing "Before" to "Vowel_Before"
    data.rename(columns=lambda x: x.replace('Before', 'Preceding_Vowel'), inplace=True)
    # Rename columns containing "After" to "Vowel_After"
    data.rename(columns=lambda x: x.replace('After', 'Following_Vowel'), inplace=True)

def plot_comparison_graph(data_before, data_after, column_names, exclude_non_sounds=False, graph_type='line'):
    plt.figure(figsize=(10, 6))

    for column_name in column_names:
        # Optionally filter data to exclude sounds with "non-" before them
        if exclude_non_sounds:
            filtered_data_before = data_before[~data_before['Sound'].str.startswith('non-')]
            filtered_data_after = data_after[~data_after['Sound'].str.startswith('non-')]
        else:
            filtered_data_before = data_before
            filtered_data_after = data_after

        # Group filtered data by sound category and calculate the mean for each category
        grouped_before = filtered_data_before.groupby('Sound')[column_name].mean()
        grouped_after = filtered_data_after.groupby('Sound')[column_name].mean()

        # Plot data based on the specified graph type
        if graph_type == 'line':
            plt.plot(grouped_before, label=f'{column_name} Before', linestyle='-', marker='o')
            plt.plot(grouped_after, label=f'{column_name} After', linestyle='--', marker='o')
        elif graph_type == 'bar':
            plt.bar(grouped_before.index, grouped_before, label=f'{column_name} Before', alpha=0.5)
            plt.bar(grouped_after.index, grouped_after, label=f'{column_name} After', alpha=0.5)
        elif graph_type == 'double_bar':
            x = np.arange(len(grouped_before))
            width = 0.35
            plt.bar(x - width/2, grouped_before, width, label=f'{column_name} Before')
            plt.bar(x + width/2, grouped_after, width, label=f'{column_name} After')

    plt.xlabel('Sound')
    plt.ylabel('Average Frequency (Hz)')
    plt.title('Average Frequency Before and After Priming')
    plt.xticks(rotation=45)
    plt.legend()
    plt.grid(True)
    plt.show()



# Example usage
file_path = 'consolidated_data.xlsx'
data_before = pd.read_excel(file_path, sheet_name='Before_Priming')  # Assuming Sheet1 contains data before priming
data_after = pd.read_excel(file_path, sheet_name='After_Priming')   # Assuming Sheet2 contains data after priming

# Define the columns you want to clean
columns_to_clean = ['Time_s_Before', 'Time_s_After', 'F1_Hz_Before', 'F2_Hz_Before', 'F3_Hz_Before', 'F4_Hz_Before',
                    'F1_Hz_After', 'F2_Hz_After', 'F3_Hz_After', 'F4_Hz_After']

# Clean non-numeric values in specified columns
clean_numeric_columns(data_before, columns_to_clean)
clean_numeric_columns(data_after, columns_to_clean)

# Rename columns containing "Before" or "After"
rename_columns(data_before)
rename_columns(data_after)

print(data_before.columns)
print(data_after.columns)

# Plot comparison graphs for both 4-line and 2-line graphs
plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel'], True, 'line')
plot_comparison_graph(data_before, data_after, ['F3_Hz_Following_Vowel'], True, 'line')
plot_comparison_graph(data_before, data_after, ['F4_Hz_Preceding_Vowel'], True, 'line')
plot_comparison_graph(data_before, data_after, ['F4_Hz_Following_Vowel'], True, 'line')

plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel', 'F3_Hz_Following_Vowel'])
plot_comparison_graph(data_before, data_after, ['F4_Hz_Preceding_Vowel', 'F4_Hz_Following_Vowel'])
plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel', 'F4_Hz_Preceding_Vowel'])
plot_comparison_graph(data_before, data_after, ['F3_Hz_Following_Vowel', 'F4_Hz_Following_Vowel'])

# plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel', 'F2_Hz_Preceding_Vowel', 'F3_Hz_Following_Vowel', 'F2_Hz_Following_Vowel'])
# plot_comparison_graph(data_before, data_after, ['F4_Hz_Preceding_Vowel', 'F2_Hz_Preceding_Vowel', 'F4_Hz_Following_Vowel', 'F2_Hz_Following_Vowel'])

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------

import matplotlib.pyplot as plt

# ANOVA results
results = {
    "F3_Hz_Before": {
        "F-statistic": 5.436721657013648,
        "P-value": 0.01985991765853451,
        "Useful result?": "Yes",
        "Conclusion": "There is a significant difference between groups."
    },
    "F3_Hz_After": {
        "F-statistic": 3.4555989534872498,
        "P-value": 0.06324925545170627,
        "Useful result?": "No",
        "Conclusion": "There is no significant difference between groups."
    },
    "F4_Hz_Before": {
        "F-statistic": 2.8194753516734523,
        "P-value": 0.09335207117691748,
        "Useful result?": "No",
        "Conclusion": "There is no significant difference between groups."
    },
    "F4_Hz_After": {
        "F-statistic": 2.5978292895413744,
        "P-value": 0.10723737364934668,
        "Useful result?": "No",
        "Conclusion": "There is no significant difference between groups."
    }
}

# Extracting data for plotting
columns = list(results.keys())
p_values = [results[col]["P-value"] for col in columns]
conclusions = [results[col]["Conclusion"] for col in columns]

# Plotting
plt.figure(figsize=(10, 6))
plt.barh(columns, p_values, color='skyblue')
plt.xlabel('P-value')
plt.title('P-values from ANOVA Results')
plt.xlim(0, 0.15)  # Adjust the x-axis limit if needed
plt.gca().invert_yaxis()  # Invert y-axis to display the highest P-value at the top
plt.grid(axis='x', linestyle='--', alpha=0.7)
plt.axvline(x=0.05, color='red', linestyle='--', linewidth=1.5, label='Significance Threshold (0.05)')
plt.axvline(x=0.1, color='orange', linestyle='--', linewidth=1.5, label='Significance Threshold (0.1)')
plt.legend()
plt.tight_layout()

# Displaying conclusions
for i, col in enumerate(columns):
    plt.text(p_values[i] + 0.001, i, conclusions[i], va='center')

plt.show()
