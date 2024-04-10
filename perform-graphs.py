import pandas as pd
import matplotlib.pyplot as plt

def clean_numeric_columns(data, columns_to_clean):
    # Iterate through each column
    for col in columns_to_clean:
        # Convert column to numeric, coerce non-numeric values to NaN
        data[col] = pd.to_numeric(data[col], errors='coerce')

def rename_columns(data):
    # Rename columns containing "Before" to "Vowel_Before"
    data.rename(columns=lambda x: x.replace('Before', 'Preceding_Vowel'), inplace=True)
    # Rename columns containing "After" to "Vowel_After"
    data.rename(columns=lambda x: x.replace('After', 'Proceeding_Vowel'), inplace=True)

def plot_comparison_graph(data_before, data_after, column_names):
    plt.figure(figsize=(10, 6))

    for column_name in column_names:
        # Group data by sound category and calculate the mean for each category
        grouped_before = data_before.groupby('Sound')[column_name].mean()
        grouped_after = data_after.groupby('Sound')[column_name].mean()

        # Plot data before and after priming with different labels
        plt.plot(grouped_before, label=f'{column_name} Before', linestyle='-', marker='o')
        plt.plot(grouped_after, label=f'{column_name} After', linestyle='--', marker='o')

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

# Plot comparison graphs for each set of columns
plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel', 'F3_Hz_Proceeding_Vowel'])
plot_comparison_graph(data_before, data_after, ['F4_Hz_Preceding_Vowel', 'F4_Hz_Proceeding_Vowel'])
plot_comparison_graph(data_before, data_after, ['F3_Hz_Preceding_Vowel', 'F4_Hz_Preceding_Vowel'])
plot_comparison_graph(data_before, data_after, ['F3_Hz_Proceeding_Vowel', 'F4_Hz_Proceeding_Vowel'])
