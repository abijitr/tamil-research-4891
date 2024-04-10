import pandas as pd
from pymer4.models import Lmer
import rpy2.robjects.packages as rpackages
from rpy2.robjects.vectors import StrVector

# Define the package names to be installed
package_names = ["lmerTest"]

# Install the packages
utils = rpackages.importr("utils")
utils.chooseCRANmirror(ind=1)  # Select the CRAN mirror
utils.install_packages(StrVector(package_names))

# Step 1: Read the data from the Excel sheet
data = pd.read_excel("summarized-formant-data.xlsx", 2)

# Step 2: Print the data to inspect its structure
print(data)

# Step 3: Define the linear mixed effects regression model
model_formula = "F1_Hz_After + F2_Hz_After + F3_Hz_After + F4_Hz_After ~ F1_Hz_Before + F2_Hz_Before + F3_Hz_Before + F4_Hz_Before + (1|Term)"

# Step 4: Create the Lmer model object
model = Lmer(model_formula, data=data)

# Step 5: Fit the model
model.fit()

# Step 6: Print the summary of the model
print(model.summary())