import pandas as pd
import numpy as np
import os
import joblib
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import classification_report, accuracy_score, confusion_matrix, ConfusionMatrixDisplay
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# ================= CONFIGURATION =================
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tracking Data")
MODEL_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models")
os.makedirs(MODEL_DIR, exist_ok=True)

plt.style.use('seaborn-v0_8-whitegrid')


def find_latest_excel():
    """Locate the most recent Tracking_Analysis.xlsx file."""
    if not os.path.exists(DATA_DIR):
        raise FileNotFoundError(f"Data directory not found: {DATA_DIR}")
    
    files = [
        f for f in os.listdir(DATA_DIR) 
        if f.endswith('.xlsx') and 'Analysis' in f and not f.startswith('~$')
    ]
    if not files:
        raise FileNotFoundError("No Tracking_Analysis Excel file found. Run the analyzer first.")
    
    latest = max(files, key=lambda x: os.path.getmtime(os.path.join(DATA_DIR, x)))
    return os.path.join(DATA_DIR, latest)


def prepare_data(excel_path):
    """Clean data and engineer features for ML."""
    print("📥 Loading data...")
    df = pd.read_excel(excel_path, sheet_name='Tracking Data')
    
    # Only use completed deliveries for training
    df_train = df[df['delivery_status'] == 'Delivered'].copy()
    
    if len(df_train) < 15:
        print(f"⚠️  Only {len(df_train)} completed deliveries found.")
        print("📌 ML models need at least 15-20 historical records to learn patterns.")
        return None, None, None
    
    # Target: 1 = Delayed, 0 = On-Time
    df_train['target'] = (df_train['is_delayed'] == 'Yes').astype(int)
    
    # Feature Engineering
    df_train['order_dt'] = pd.to_datetime(df_train['order_confirmed'], errors='coerce')
    df_train['day_of_week'] = df_train['order_dt'].dt.dayofweek  # 0=Mon, 6=Sun
    df_train['month'] = df_train['order_dt'].dt.month
    df_train['logistics_count'] = pd.to_numeric(df_train['logistics_count'], errors='coerce').fillna(0)
    
    # Encode city
    le_city = LabelEncoder()
    df_train['city_encoded'] = le_city.fit_transform(df_train['receive_city'])
    
    feature_cols = ['day_of_week', 'month', 'logistics_count', 'city_encoded']
    X = df_train[feature_cols]
    y = df_train['target']
    
    print(f"✅ Prepared {len(X)} samples for training.")
    return X, y, le_city


def train_and_evaluate(X, y, le_city):
    """Train Random Forest and evaluate performance."""
    print("🤖 Training Random Forest Classifier...")
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y
    )
    
    model = RandomForestClassifier(
        n_estimators=150,
        max_depth=6,
        class_weight='balanced',
        random_state=42
    )
    model.fit(X_train, y_train)
    
    # Evaluate
    y_pred = model.predict(X_test)
    print("\n📊 MODEL PERFORMANCE:")
    print(classification_report(y_test, y_pred, target_names=['On-Time', 'Delayed']))
    
    # Confusion Matrix Plot
    plt.figure(figsize=(6, 5))
    cm = confusion_matrix(y_test, y_pred)
    disp = ConfusionMatrixDisplay(confusion_matrix=cm, display_labels=['On-Time', 'Delayed'])
    disp.plot(cmap='Blues', values_format='d')
    plt.title("Confusion Matrix")
    plt.savefig(os.path.join(MODEL_DIR, "confusion_matrix.png"), dpi=300)
    plt.close()
    
    # Feature Importance
    plt.figure(figsize=(8, 6))
    importances = model.feature_importances_
    features = X.columns
    indices = np.argsort(importances)[::-1]
    plt.barh(range(len(indices)), importances[indices], align='center', color='#4dabf7')
    plt.yticks(range(len(indices)), [features[i] for i in indices])
    plt.xlabel('Relative Importance')
    plt.title('Feature Importance')
    plt.gca().invert_yaxis()
    plt.tight_layout()
    plt.savefig(os.path.join(MODEL_DIR, "feature_importance.png"), dpi=300)
    plt.close()
    
    # Save model & encoder
    model_path = os.path.join(MODEL_DIR, "delivery_predictor.pkl")
    encoder_path = os.path.join(MODEL_DIR, "city_encoder.pkl")
    joblib.dump(model, model_path)
    joblib.dump(le_city, encoder_path)
    
    print(f"\n✅ Model saved to: {model_path}")
    return model


def predict_delivery(city, order_date_str, logistics_count):
    """Predict delivery status for a new shipment."""
    model_path = os.path.join(MODEL_DIR, "delivery_predictor.pkl")
    encoder_path = os.path.join(MODEL_DIR, "city_encoder.pkl")
    
    if not os.path.exists(model_path):
        raise FileNotFoundError("Model not found. Run training first.")
    
    model = joblib.load(model_path)
    le_city = joblib.load(encoder_path)
    
    # Parse date
    order_dt = pd.to_datetime(order_date_str)
    day_of_week = order_dt.dayofweek
    month = order_dt.month
    
    # Handle unseen cities gracefully
    try:
        city_enc = le_city.transform([city])[0]
    except ValueError:
        city_enc = 0  # Fallback to most common/unknown city code
    
    features = pd.DataFrame({
        'day_of_week': [day_of_week],
        'month': [month],
        'logistics_count': [logistics_count],
        'city_encoded': [city_enc]
    })
    
    pred = model.predict(features)[0]
    proba = model.predict_proba(features)[0]
    
    result = "⚠️  DELAYED" if pred == 1 else "✅ ON-TIME"
    confidence = proba[pred] * 100
    
    print("\n" + "="*40)
    print("🔮 DELIVERY PREDICTION")
    print("="*40)
    print(f"📦 City: {city}")
    print(f"📅 Order Date: {order_date_str}")
    print(f"🔄 Logistics Stops: {logistics_count}")
    print(f"🎯 Prediction: {result}")
    print(f"📊 Confidence: {confidence:.1f}%")
    print(f"📉 Prob On-Time: {proba[0]*100:.1f}% | 📈 Prob Delayed: {proba[1]*100:.1f}%")
    print("="*40)
    
    return result, confidence


# ================= MAIN EXECUTION =================
if __name__ == "__main__":
    import sys
    
    print("📈 BuffaloEx Delivery Predictor")
    print("="*40)
    
    # Mode 1: Train Model
    if len(sys.argv) > 1 and sys.argv[1] == "--train":
        try:
            excel_path = find_latest_excel()
            X, y, le_city = prepare_data(excel_path)
            if X is not None:
                train_and_evaluate(X, y, le_city)
        except Exception as e:
            print(f"❌ Training failed: {e}")
    
    # Mode 2: Predict (Interactive)
    else:
        if not os.path.exists(os.path.join(MODEL_DIR, "delivery_predictor.pkl")):
            print("⚠️  Model not trained yet. Run: python delivery_predictor.py --train")
            exit(1)
            
        print("\n🔍 Enter shipment details for prediction:")
        city = input("🏙️  Destination City: ").strip()
        date_str = input("📅 Order Date (YYYY-MM-DD): ").strip()
        stops = input("🔄 Expected Logistics Stops: ").strip()
        
        try:
            predict_delivery(city, date_str, int(stops))
        except Exception as e:
            print(f"❌ Prediction failed: {e}")