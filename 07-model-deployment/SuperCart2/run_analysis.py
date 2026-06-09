"""
Faithful replication of Low_Code_SuperKart_Model_Deployment_Notebook_[Updated]_version_2.ipynb
modeling pipeline, to generate real EDA insights + model R2/RMSE numbers.
Skips all Docker/API/deployment cells per request.
"""
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.compose import make_column_transformer
from sklearn.pipeline import make_pipeline, Pipeline
from sklearn.preprocessing import OneHotEncoder
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import (
    BaggingRegressor,
    RandomForestRegressor,
    AdaBoostRegressor,
    GradientBoostingRegressor,
)
from sklearn.metrics import r2_score, mean_squared_error, mean_absolute_error
from xgboost import XGBRegressor

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

# ----------------------------------------------------------------------
# 1. Load + clean (mirrors notebook)
# ----------------------------------------------------------------------
kart = pd.read_csv("SuperKart.csv")
data = kart.copy()

print("=" * 70)
print("DATASET OVERVIEW")
print("=" * 70)
print(f"Rows: {data.shape[0]}, Columns: {data.shape[1]}")
print(f"Duplicate rows: {data.duplicated().sum()}")
print(f"Missing values total: {data.isnull().sum().sum()}")

# Fix sugar content label (reg -> Regular)
data.Product_Sugar_Content.replace(to_replace=["reg"], value=["Regular"], inplace=True)

# ----------------------------------------------------------------------
# 2. Feature engineering (mirrors notebook)
# ----------------------------------------------------------------------
data["Product_Id_char"] = data["Product_Id"].str[:2]
data["Store_Age_Years"] = 2025 - data.Store_Establishment_Year

perishables = ["Dairy", "Meat", "Fruits and Vegetables", "Breakfast", "Breads", "Seafood"]
def change(x):
    return "Perishables" if x in perishables else "Non Perishables"
data["Product_Type_Category"] = data["Product_Type"].apply(change)

# ----------------------------------------------------------------------
# 3. Quick EDA insights
# ----------------------------------------------------------------------
print("\n" + "=" * 70)
print("EDA INSIGHTS")
print("=" * 70)
tgt = "Product_Store_Sales_Total"
print(f"\nTarget ({tgt}) summary:")
print(data[tgt].describe().round(2).to_string())

print("\nCorrelation of numeric features with target:")
num_cols = data.select_dtypes(include=np.number).columns.tolist()
corr = data[num_cols].corr()[tgt].drop(tgt).sort_values(ascending=False)
print(corr.round(3).to_string())

print("\nMean sales by Store_Type:")
print(data.groupby("Store_Type")[tgt].mean().round(0).sort_values(ascending=False).to_string())
print("\nMean sales by Store_Size:")
print(data.groupby("Store_Size")[tgt].mean().round(0).sort_values(ascending=False).to_string())
print("\nMean sales by Store_Location_City_Type:")
print(data.groupby("Store_Location_City_Type")[tgt].mean().round(0).sort_values(ascending=False).to_string())
print("\nTotal revenue by Store_Id:")
print(data.groupby("Store_Id")[tgt].sum().round(0).sort_values(ascending=False).to_string())

# ----------------------------------------------------------------------
# 4. Drop unneeded cols, split
# ----------------------------------------------------------------------
data = data.drop(["Product_Id", "Product_Type", "Store_Id", "Store_Establishment_Year"], axis=1)

X = data.drop(tgt, axis=1)
y = data[tgt]
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=1, shuffle=True
)
print("\n" + "=" * 70)
print(f"Train/Test split -> train: {X_train.shape}, test: {X_test.shape}")
print("=" * 70)

# Preprocessor: OHE categoricals, passthrough numerics (remainder='passthrough'
# so the strong numeric predictors like Product_MRP are NOT dropped).
categorical_features = X.select_dtypes(include=["object", "category"]).columns.tolist()
preprocessor = make_column_transformer(
    (Pipeline([("encoder", OneHotEncoder(handle_unknown="ignore"))]), categorical_features),
    remainder="passthrough",
)

# ----------------------------------------------------------------------
# 5. Performance helper (mirrors notebook)
# ----------------------------------------------------------------------
def adj_r2(predictors, targets, predictions):
    r2 = r2_score(targets, predictions)
    n = predictors.shape[0]
    k = predictors.shape[1]
    return 1 - ((1 - r2) * (n - 1) / (n - k - 1))

def perf(model, Xp, yt):
    pred = model.predict(Xp)
    return {
        "RMSE": np.sqrt(mean_squared_error(yt, pred)),
        "MAE": mean_absolute_error(yt, pred),
        "R2": r2_score(yt, pred),
        "AdjR2": adj_r2(Xp, yt, pred),
        "MAPE": np.mean(np.abs((yt - pred) / yt)),
    }

# ----------------------------------------------------------------------
# 6. Base models
# ----------------------------------------------------------------------
base_models = {
    "Decision Tree": DecisionTreeRegressor(random_state=1),
    "Bagging": BaggingRegressor(random_state=1),
    "Random Forest": RandomForestRegressor(random_state=1),
    "AdaBoost": AdaBoostRegressor(random_state=1),
    "Gradient Boosting": GradientBoostingRegressor(random_state=1),
    "XGBoost": XGBRegressor(random_state=1, verbosity=0),
}

train_rows, test_rows = [], []
for name, est in base_models.items():
    pipe = make_pipeline(preprocessor, est)
    pipe.fit(X_train, y_train)
    tr, te = perf(pipe, X_train, y_train), perf(pipe, X_test, y_test)
    train_rows.append({"Model": name, **tr})
    test_rows.append({"Model": name, **te})
    print(f"  trained base: {name}")

train_df = pd.DataFrame(train_rows).set_index("Model").round(3)
test_df = pd.DataFrame(test_rows).set_index("Model").round(3)

print("\n" + "=" * 70)
print("BASE MODELS - TRAINING PERFORMANCE")
print("=" * 70)
print(train_df.to_string())
print("\n" + "=" * 70)
print("BASE MODELS - TEST PERFORMANCE")
print("=" * 70)
print(test_df.to_string())

# ----------------------------------------------------------------------
# 7. Hyperparameter tuning for the two finalists: Random Forest & XGBoost
# ----------------------------------------------------------------------
print("\n" + "=" * 70)
print("HYPERPARAMETER TUNING (GridSearchCV, cv=3, scoring='r2')")
print("=" * 70)

tuned_results = {}

# Random Forest
rf_pipe = make_pipeline(preprocessor, RandomForestRegressor(random_state=1))
rf_params = {
    "randomforestregressor__n_estimators": [100, 200, 300],
    "randomforestregressor__max_depth": [7, 10, 15, None],
    "randomforestregressor__max_features": ["sqrt", 0.5, 0.8],
}
rf_grid = GridSearchCV(rf_pipe, rf_params, scoring="r2", cv=3, n_jobs=-1)
rf_grid.fit(X_train, y_train)
print("\nRandom Forest best params:", rf_grid.best_params_)
tuned_results["Random Forest (tuned)"] = perf(rf_grid.best_estimator_, X_test, y_test)

# XGBoost
xgb_pipe = make_pipeline(preprocessor, XGBRegressor(random_state=1, verbosity=0))
xgb_params = {
    "xgbregressor__n_estimators": [100, 200],
    "xgbregressor__subsample": [0.8, 1.0],
    "xgbregressor__gamma": [0, 5],
    "xgbregressor__colsample_bytree": [0.8, 1.0],
    "xgbregressor__colsample_bylevel": [0.8, 1.0],
}
xgb_grid = GridSearchCV(xgb_pipe, xgb_params, scoring="r2", cv=3, n_jobs=-1)
xgb_grid.fit(X_train, y_train)
print("XGBoost best params:", xgb_grid.best_params_)
tuned_results["XGBoost (tuned)"] = perf(xgb_grid.best_estimator_, X_test, y_test)

tuned_df = pd.DataFrame(tuned_results).T.round(3)
print("\n" + "=" * 70)
print("TUNED MODELS - TEST PERFORMANCE")
print("=" * 70)
print(tuned_df.to_string())

# ----------------------------------------------------------------------
# 8. Final leaderboard (test R2)
# ----------------------------------------------------------------------
leaderboard = pd.concat([test_df, tuned_df]).sort_values("R2", ascending=False)
print("\n" + "=" * 70)
print("FINAL LEADERBOARD (sorted by Test R2)")
print("=" * 70)
print(leaderboard[["R2", "AdjR2", "RMSE", "MAE", "MAPE"]].to_string())
