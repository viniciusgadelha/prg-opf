import numpy as np
from calc_mmc_losses import *
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline
import matplotlib.pyplot as plt

# Generate sample data
num_samples = 50000
I_values = np.random.uniform(low=0, high=0.1, size=num_samples)  # Example range for I^2
V_values = np.random.uniform(low=0, high=50, size=num_samples)  # Example range for V^2 as a base of 6kV
outputs = [calc_mmc_losses(current, voltage) for current, voltage in zip(I_values, V_values)]

X = np.column_stack((I_values, V_values))
X_train, X_test, y_train, y_test = train_test_split(X, outputs, test_size=0.2, random_state=42)

# Create and train the model
model = LinearRegression()
model.fit(X_train, y_train)

# Make predictions on the test set
y_pred = model.predict(X_test)

# Calculate metrics
mse = mean_squared_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)
print("Mean Squared Error:", mse)
print("R-squared:", r2)

coefficients = model.coef_
intercept = model.intercept_

print("Coefficient for I:", coefficients[0])
print("Coefficient for V:", coefficients[1])
print("Intercept:", intercept)

new_I = 0.027889
new_V = 36
predicted_output = model.predict([[new_I, new_V]])
print("Predicted output for I={}, V={}: {}".format(new_I, new_V, predicted_output[0]))

# --------------------------------------------------------------------------------------------------------------


# Generate sample data
num_samples = 500000
P_values = np.random.uniform(low=0, high=2.5, size=num_samples)  # Example range for P (0-5 MW)
outputs = [calc_mmc_losses(p_mw=power) for power in P_values]

X_train, X_test, y_train, y_test = train_test_split(P_values, outputs, test_size=0.2, random_state=42)

# Create and train the model
model = LinearRegression()
model.fit(X_train.reshape(-1, 1), y_train)

# Make predictions on the test set
y_pred = model.predict(X_test.reshape(-1, 1))

# Calculate metrics
mse = mean_squared_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)
print("Mean Squared Error:", mse)
print("R-squared:", r2)

coefficients = model.coef_
intercept = model.intercept_

print("Coefficient for P:", coefficients[0])
print("Intercept:", intercept)

new_P = 1
new_V = 36
predicted_output = model.predict([[new_P]])
print("Predicted output for P={} MW: {} W".format(new_P, predicted_output[0]*1000000))



# -----------------------------------------------------------------------------------------------------------------



P_range = [-50, 50]
freq_range = [100, 10000]
V_range = [25, 49]
P = 1
V = 36

freq_values = np.linspace(freq_range[0], freq_range[1], 100)

from calc_mmc_losses import *
import matplotlib.pyplot as plt
import pandas as pd

P_values = np.linspace(-2.5, 2.5, 100)
# losses = [calc_mmc_losses(P, V) for P in P_values]
losses = [calc_mmc_losses(p_mw=f) for f in P_values]


losses = [x * 1000000 for x in losses]

losses_df = pd.Series(losses)
losses_df.to_excel('1-2.xlsx')

# Plot
plt.figure(figsize=(10, 6))
# plt.plot(P_values, losses, marker='o')
plt.plot(P_values, losses, marker='')
plt.xlabel('P MW')
plt.ylabel('LOSSES W')
# plt.title('MMC Losses (V = {})'.format(V))
plt.grid(True)
plt.show()
