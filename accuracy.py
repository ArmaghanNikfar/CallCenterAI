import pandas as pd
import numpy as np
from sklearn.metrics import mean_squared_error, mean_absolute_error
from keras._tf_keras.keras.models import load_model
import matplotlib.pyplot as plt

# Read predicted and actual calls from Excel files
try:
    predicted_data = pd.read_excel('Finnal1403BLSTMPrediction.xlsx')
    real_data = pd.read_excel('Incoming1402pure.xlsx')
    model = load_model('prediction_model.h5')

    # Ensure the data is sorted and aligned properly by date and hour if necessary
    predicted_data = predicted_data.sort_values(by=['Date', 'Hours'])
    real_data = real_data.sort_values(by=['Date', 'Hours'])
    hours = real_data['Hours'].values
    # Extract the calls columns
    predicted_calls = predicted_data['Predicted Calls'].values
    real_calls = real_data['Calls'].values


    # Check for zero values in real_calls and handle them
    mask = real_calls != 0
    filtered_real_calls = real_calls[mask]
    filtered_predicted_calls = predicted_calls[mask]

    # Calculate total actual and predicted calls
    total_real_calls = np.sum(filtered_real_calls)
    total_predicted_calls = np.sum(filtered_predicted_calls)

    # Calculate overall percentage error
    overall_percentage_error = np.abs((total_real_calls - total_predicted_calls) / total_real_calls) * 100

    # Calculate accuracy based on a threshold (e.g., 10% of the actual value)
    threshold = 10
    correct_predictions = np.abs(real_calls - predicted_calls) <= (threshold * real_calls)
    accuracy = np.mean(correct_predictions) * 100

    # Print overall percentage error and accuracy
    print(f"Overall Percentage Error: {overall_percentage_error:.2f}%")
    print(f"Accuracy: {accuracy:.2f}%")


 # Plot the predicted and real calls
    plt.figure(figsize=(14, 7))
    plt.plot(predicted_calls, label='Predicted Calls', color='blue',linestyle='-',linewidth=1)
    plt.plot(real_calls, label='Real Calls', color='green',linestyle='-',linewidth=1)
    plt.xticks([])  # Hide x-axis values
    plt.yticks([])  # Hide y-axis values
    plt.title('Predicted vs Real Calls')
    plt.legend()
    plt.show()

except Exception as e:
    print(f"Error processing Excel files: {e}")
