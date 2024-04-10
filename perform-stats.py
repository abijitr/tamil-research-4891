import pandas as pd
import numpy as np
import statsmodels.api as sm
import statsmodels.formula.api as smf

def clean_numeric_columns(data):
    # Iterate through each column
    for col in data.columns:
        # Convert column to numeric, coerce non-numeric values to NaN
        data[col] = pd.to_numeric(data[col], errors='coerce')

    # Fill NaN values with 0
    data.fillna(0, inplace=True)

def perform_statistical_analysis(file_path):
    # Load data from the Excel file
    data_before = pd.read_excel(file_path, sheet_name='Before_Priming')
    data_after = pd.read_excel(file_path, sheet_name='After_Priming')

    # Clean non-numeric values in all columns
    clean_numeric_columns(data_before)
    clean_numeric_columns(data_after)

    # Fit statistical models for F3_Hz
    model1_f3 = smf.ols('F3_Hz_After ~ F3_Hz_Before + Time_s_Before + Time_s_After', data=data_before)
    results1_f3 = model1_f3.fit()

    model2_f3 = smf.ols('F3_Hz_After ~ F3_Hz_Before + Time_s_Before + Time_s_After', data=data_after)
    results2_f3 = model2_f3.fit()

    # Fit statistical models for F4_Hz
    model1_f4 = smf.ols('F4_Hz_After ~ F4_Hz_Before + Time_s_Before + Time_s_After', data=data_before)
    results1_f4 = model1_f4.fit()

    model2_f4 = smf.ols('F4_Hz_After ~ F4_Hz_Before + Time_s_Before + Time_s_After', data=data_after)
    results2_f4 = model2_f4.fit()

    return (results1_f3, results2_f3, results1_f4, results2_f4)


# Example usage
file_path = 'consolidated_data.xlsx'
results_before_f3, results_after_f3, results_before_f4, results_after_f4 = perform_statistical_analysis(file_path)
print("Results for f3 data before priming:")
print(results_before_f3.summary())
print("\nResults for f3 data after priming:")
print(results_after_f3.summary())
print("Results for f4 data before priming:")
print(results_before_f4.summary())
print("\nResults for f4 data after priming:")
print(results_after_f4.summary())