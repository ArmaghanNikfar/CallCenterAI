import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from keras.api.models import Sequential
from keras.api.layers import Dense
from keras.api.layers import LSTM, Bidirectional
from keras.api.models import load_model

def train_and_save_model(dates, hours, calls, is_ramadan, is_eid, is_moharam , is_national_holiday , is_day_off):
    features = np.column_stack((hours, calls, is_ramadan, is_eid , is_moharam , is_national_holiday , is_day_off)).astype(np.float32)
    features_train, features_test, calls_train, calls_test = train_test_split(features, calls, test_size=0.2, random_state=42)
    calls_train = calls_train.astype(np.float32)

    model = Sequential([
        Bidirectional(LSTM(64, return_sequences=True), input_shape=(1, 7)),  # Bidirectional LSTM with return_sequences
        Bidirectional(LSTM(32)),
        Dense(30, activation='relu'),
        Dense(50, activation='relu'),
        Dense(90, activation='relu'),
        Dense(150, activation='relu'),
        Dense(200, activation='relu'),
        Dense(1, activation='linear')
    ])

    model.compile(optimizer='adam', loss='mean_squared_error')
    model.fit(features_train.reshape(features_train.shape[0], 1, features_train.shape[1]), calls_train, epochs=10, verbose=0, batch_size=100)  # Reshape for sequence input

    # Save the model
    model.save('prediction_model.h5')

# Read data from excel
data = pd.read_excel('AllCalls.xlsx')
hours = data['Hours'].values
calls = data['Calls'].values
dates = data['Date'].values
is_ramadan = data['IsRamadan'].values
is_eid = data['IsEidNowrooz'].values
is_moharam = data['IsMoharam'].values
is_national_holiday = data['IsNationalHoliday'].values
is_day_off = data['ItsDayOff'].values
# Train and save the model
train_and_save_model(dates, hours, calls, is_ramadan, is_eid, is_moharam , is_national_holiday,is_day_off)
#Negin is online
print("Negin")