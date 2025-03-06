import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from keras.api.models import Sequential
from keras.api.layers import Dense, LSTM, Bidirectional, Dropout
from keras.api.callbacks import EarlyStopping, ReduceLROnPlateau

def train_and_save_model(dates, hours, calls, is_ramadan, is_eid, is_moharam, is_national_holiday, is_day_off):
    features = np.column_stack((hours, is_ramadan, is_eid, is_moharam, is_national_holiday, is_day_off)).astype(np.float32)

    # Normalize the features
    scaler = StandardScaler()
    features = scaler.fit_transform(features)

    features_train, features_test, calls_train, calls_test = train_test_split(features, calls, test_size=0.2, random_state=42)
    calls_train = calls_train.astype(np.float32)

    model = Sequential([
        Bidirectional(LSTM(64, return_sequences=True), input_shape=(1,7)),
        Dropout(0.2),
        Bidirectional(LSTM(32)),
        Dropout(0.2),
        Dense(64, activation='relu'),
        Dense(32, activation='relu'),
        Dense(16, activation='relu'),
        Dense(1, activation='linear')
    ])

    model.compile(optimizer='adam', loss='mean_squared_error', metrics=['mae'])

    # Callbacks
    early_stopping = EarlyStopping(monitor='val_loss', patience=10, restore_best_weights=True)
    reduce_lr = ReduceLROnPlateau(monitor='val_loss', factor=0.5, patience=5, min_lr=1e-5)

    model.fit(
        features_train.reshape(features_train.shape[0], 1, features_train.shape[1]),
        calls_train,
        validation_split=0.2,
        epochs=100,
        batch_size=32,
        verbose=1,
        callbacks=[early_stopping, reduce_lr]
    )

    # Save the model
    model.save('prediction_BLSTMmodel.h5')

    return model, scaler

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
model, scaler = train_and_save_model(dates, hours, calls, is_ramadan, is_eid, is_moharam, is_national_holiday, is_day_off)

# Save the scaler for future use
import joblib
joblib.dump(scaler, 'scaler.pkl')
