import numpy as np
import pandas as pd
from keras._tf_keras.keras.models import load_model

# Read new data from Excel
try:
    new_data = pd.read_excel('1402To1403ChangeyearAndSetFeatures.xlsx')
    new_hours = new_data['Hours'].values
    new_dates = new_data['Date'].values
    new_calls = new_data['Predicted Calls'].values
    new_is_ramadan = new_data['Isramadan'].values
    new_is_eid = new_data['Isnowruz'].values
    new_is_moharam = new_data['Ismoharam'].values
    new_in_nationalHoliday = new_data['Isnationalholiday'].values
    new_is_day_off = new_data['Isreligiousholidays'].values
    
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

def calculate_agents(call_volume):
    avg_call_duration = 7  # Average call duration in minutes
    calls_per_hour_per_agent = 60 / avg_call_duration

    if call_volume < 10:
        max_logoff_agents = 3
    elif 10 <= call_volume <= 20:
        max_logoff_agents = 5
    elif 20 < call_volume <= 30:
        max_logoff_agents = 7
    else:
        max_logoff_agents = 14

    # Calculate required agents based on call volume
    active_agents_required = np.ceil(call_volume / calls_per_hour_per_agent)
    total_required_agents = active_agents_required + max_logoff_agents

    return int(total_required_agents)

def predict_calls_and_agents(dates, hours, calls, is_ramadan, is_eid, is_moharam, is_national_holiday, is_day_off, model_path='prediction_BLSTMmodel.h5'):
    hours_excel = []
    agencies_excel = []
    predicted_calls_list = []

    try:
       
        model = load_model(model_path)
    except Exception as e:
        print(f"Error loading model: {e}")
        return pd.DataFrame()

    for date, hour, call, ramadan, eid, moharam, national_holiday, day_off in zip(dates, hours, calls, is_ramadan, is_eid, is_moharam, is_national_holiday,  is_day_off):
     
        prediction_input = np.array([[hour, call, ramadan, eid, moharam, national_holiday, day_off]]).astype(np.float32)
        prediction_input = prediction_input.reshape(1, 1, prediction_input.shape[1])  # Reshape for sequence input

        # Predict using the trained model
        try:
            predicted_call = model.predict(prediction_input, verbose=0)
            predicted_call = int(predicted_call[0, 0])
        except Exception as e:
            print(f"Error predicting calls: {e}")
            predicted_call = 0

        predicted_calls_list.append(predicted_call)
        
        # Calculate required agents based on predicted calls
        predicted_agents = calculate_agents(predicted_call)

        print(f"Date: {date}, Hour: {hour}, Predicted agents needed: {predicted_agents}, Predicted calls: {predicted_call}")
        
        hours_excel.append(hour)
        agencies_excel.append(predicted_agents)

    result_df = pd.DataFrame({
        'Date': dates,
        'Hours': hours_excel,
        'Predicted Calls': predicted_calls_list,
        'Agencies': agencies_excel
    })
    
    return result_df

# Predict and print for new data
print("\nNew Predictions:")
new_result = predict_calls_and_agents(new_dates, new_hours, new_calls, new_is_ramadan, new_is_eid, new_is_moharam, new_in_nationalHoliday, new_is_day_off)

if not new_result.empty:
    try:
        # Export new results to Excel
        new_result.to_excel('Finnal1403BLSTMPrediction.xlsx', index=False)
        print("New predictions have been saved to 'NewPredictedAgencies.xlsx'")
    except Exception as e:
        print(f"Error saving Excel file: {e}")
else:
    print("No results to save.")
