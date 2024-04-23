import pandas as pd
import numpy as np
import statsmodels.api as sm
import statsmodels.formula.api as smf
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error

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

    # Fit scikit-learn linear regression models for F3_Hz and F4_Hz
    lr_before_f3 = LinearRegression().fit(data_before[['F3_Hz_Before', 'Time_s_Before', 'Time_s_After']], data_before['F3_Hz_After'])
    lr_after_f3 = LinearRegression().fit(data_after[['F3_Hz_Before', 'Time_s_Before', 'Time_s_After']], data_after['F3_Hz_After'])
    lr_before_f4 = LinearRegression().fit(data_before[['F4_Hz_Before', 'Time_s_Before', 'Time_s_After']], data_before['F4_Hz_After'])
    lr_after_f4 = LinearRegression().fit(data_after[['F4_Hz_Before', 'Time_s_Before', 'Time_s_After']], data_after['F4_Hz_After'])

    return data_before, data_after, results1_f3, results2_f3, results1_f4, results2_f4, lr_before_f3, lr_after_f3, lr_before_f4, lr_after_f4


def print_rmse_results(model, X_train, y_train, X_test, y_test, model_name):
    # Predictions
    y_train_pred = model.predict(X_train)
    y_test_pred = model.predict(X_test)

    # Calculate RMSE
    rmse_train = np.sqrt(mean_squared_error(y_train, y_train_pred))
    rmse_test = np.sqrt(mean_squared_error(y_test, y_test_pred))

    print(f"RMSE for {model_name} (Train): {rmse_train}")
    print(f"RMSE for {model_name} (Test): {rmse_test}")


# Example usage
file_path = 'consolidated_data.xlsx'
data_before, data_after, results_before_f3, results_after_f3, results_before_f4, results_after_f4, lr_before_f3, lr_after_f3, lr_before_f4, lr_after_f4 = perform_statistical_analysis(file_path)

# Print results for statistical models
print("OLS Results for f3 data before priming:")
print(results_before_f3.summary())
print("\nOLS Results for f3 data after priming:")
print(results_after_f3.summary())
print("OLS Results for f4 data before priming:")
print(results_before_f4.summary())
print("\nOLS Results for f4 data after priming:")
print(results_after_f4.summary())

def print_lr_summary(model):
    # Extract coefficients
    coefficients = model.coef_
    intercept = model.intercept_

    # Print summary
    print("Linear Regression Summary:")
    print("Intercept:", intercept)
    for i, coef in enumerate(coefficients):
        print(f"Coefficient {i+1}: {coef}")



# Print results for statistical models
print("LR Results for f3 data before priming:")
print(print_lr_summary(lr_before_f3))
print("\nLR Results for f3 data after priming:")
print(print_lr_summary(lr_after_f3))
print("LR Results for f4 data before priming:")
print(print_lr_summary(lr_before_f4))
print("\nLR Results for f4 data after priming:")
print(print_lr_summary(lr_after_f4))

# Define X and y for F3_Hz before and after priming
X_before_f3 = data_before[['F3_Hz_Before', 'Time_s_Before', 'Time_s_After']]
y_before_f3 = data_before['F3_Hz_After']
X_after_f3 = data_after[['F3_Hz_Before', 'Time_s_Before', 'Time_s_After']]
y_after_f3 = data_after['F3_Hz_After']

# Define X and y for F4_Hz before and after priming
X_before_f4 = data_before[['F4_Hz_Before', 'Time_s_Before', 'Time_s_After']]
y_before_f4 = data_before['F4_Hz_After']
X_after_f4 = data_after[['F4_Hz_Before', 'Time_s_Before', 'Time_s_After']]
y_after_f4 = data_after['F4_Hz_After']

print("PRINTING RMSE")

# Print RMSE for scikit-learn linear regression models
print_rmse_results(lr_before_f3, X_before_f3, y_before_f3, X_after_f3, y_after_f3, "Linear Regression F3")
print_rmse_results(lr_before_f4, X_before_f4, y_before_f4, X_after_f4, y_after_f4, "Linear Regression F4")

import subprocess

# Specify the path to the Python file you want to run
python_file_path = 'perform-graphs.py'

# Run the Python file
print('running: ', python_file_path)
# subprocess.run(['python', python_file_path])


