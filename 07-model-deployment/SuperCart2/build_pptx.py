"""
Populate 'SuperKart Presentation - Final.pptx' template with real results
from run_analysis.py. Preserves the template's layout/branding.
"""
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SRC = "SuperKart Presentation - Final.pptx"

ACCENT = RGBColor(0x1F, 0x4E, 0x79)   # dark blue
HEADER = RGBColor(0x2E, 0x75, 0xB6)   # medium blue
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xDE, 0xEA, 0xF6)
DARK = RGBColor(0x33, 0x33, 0x33)

prs = Presentation(SRC)


def set_body(tf, items, base_size=14, sub_size=12):
    """items: list of (text, level). Reuses paragraph 0, rebuilds the rest."""
    # wipe existing paragraphs except the first
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    # remove extra paragraphs
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    # clear runs in p0
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)

    first = True
    for text, level in items:
        para = p0 if first else tf.add_paragraph()
        first = False
        para.level = level
        run = para.add_run()
        run.text = text
        run.font.size = Pt(base_size if level == 0 else sub_size)
        run.font.color.rgb = DARK
        para.space_after = Pt(4)


def add_table(slide, rows, left, top, width, height, col_widths=None,
              font_size=10, header_size=10):
    nrows = len(rows)
    ncols = len(rows[0])
    gtbl = slide.shapes.add_table(nrows, ncols, Inches(left), Inches(top),
                                  Inches(width), Inches(height))
    tbl = gtbl.table
    if col_widths:
        for c, w in enumerate(col_widths):
            tbl.columns[c].width = Inches(w)
    for r in range(nrows):
        for c in range(ncols):
            cell = tbl.cell(r, c)
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            tf = cell.text_frame
            tf.word_wrap = True
            para = tf.paragraphs[0]
            run = para.add_run()
            run.text = str(rows[r][c])
            run.font.size = Pt(header_size if r == 0 else font_size)
            if r == 0:
                run.font.bold = True
                run.font.color.rgb = WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = HEADER
            else:
                run.font.color.rgb = DARK
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT if r % 2 == 0 else WHITE
            if c == 0:
                para.alignment = PP_ALIGN.LEFT
            else:
                para.alignment = PP_ALIGN.CENTER
    return gtbl


S = prs.slides

# ---------------- Slide 1: Title ----------------
s = S[0]
s.shapes[0].text_frame.paragraphs[0].runs[0].text = "SuperKart Sales Forecasting & Model Deployment"
# subtitle / course
tf = s.shapes[1].text_frame
tf.paragraphs[0].runs[0].text = "Predicting Product-Store Sales with Machine Learning  |  Model Deployment"
s.shapes[2].text_frame.paragraphs[0].runs[0].text = "June 2026"

# ---------------- Slide 3: Executive Summary ----------------
set_body(S[2].shapes[1].text_frame, [
    ("A machine-learning model was built to forecast product-level sales for each store, using historical data on 8,763 product-store records across 4 outlets.", 0),
    ("Random Forest and XGBoost (with and without tuning) were compared; all four are statistically tied at Test R² ≈ 0.668 (RMSE ≈ ₹616, MAPE ≈ 18.7%), with a slight, consistent edge to Random Forest.", 0),
    ("Tuned Random Forest was selected and deployed as the production model.", 0),
    ("Product MRP (price) is the strongest correlate of sales (r = 0.79); a controlled test shows adding it would lift Test R² from 0.67 to 0.93 — a recommended future enhancement (see Recommendation slide).", 0),
    ("Store format drives revenue: Departmental Stores and Tier-1 cities generate ~3x the sales of Food Marts and Tier-3 locations; OUT004 alone produced ₹15.4M — more than the other three outlets combined.", 0),
    ("Recommendation: operate the tuned Random Forest forecasting service (Flask API + Streamlit app) to guide inventory, supply-chain planning, and store strategy.", 0),
])

# ---------------- Slide 4: Business Problem & Approach ----------------
set_body(S[3].shapes[1].text_frame, [
    ("Business problem", 0),
    ("SuperKart needs accurate sales forecasts at the product-store level to plan inventory, optimize the supply chain, and benchmark store performance.", 1),
    ("The target, Product_Store_Sales_Total, is continuous — this is a regression problem.", 1),
    ("Solution approach / methodology", 0),
    ("Exploratory data analysis to understand sales drivers and data quality.", 1),
    ("Feature engineering + a preprocessing pipeline (one-hot encoding of categoricals).", 1),
    ("Train and compare 6 tree-based / ensemble regressors on R², RMSE, MAE, MAPE.", 1),
    ("Hyperparameter-tune the strongest candidates, select the best, and deploy it as an API + web app.", 1),
])

# ---------------- Slide 5: EDA Results ----------------
set_body(S[4].shapes[1].text_frame, [
    ("Target: average sale ₹3,464 (range ₹33–₹8,000), roughly symmetric distribution.", 0),
    ("Strongest drivers of sales: Product_MRP (corr 0.79) and Product_Weight (corr 0.74); Product_Allocated_Area shows ~no linear relationship (-0.00).", 0),
    ("Store_Type: Departmental ₹4,947 > Supermarket Type1 ₹3,924 > Supermarket Type2 ₹3,299 > Food Mart ₹1,763.", 0),
    ("City tier: Tier 1 ₹4,947 > Tier 2 ₹3,457 > Tier 3 ₹1,763; Store size: High > Medium > Small.", 0),
    ("Revenue by outlet: OUT004 ₹15.4M, OUT003 ₹6.7M, OUT001 ₹6.2M, OUT002 ₹2.0M.", 0),
    ("Implication: pricing/product attributes plus store format almost fully determine expected sales.", 0),
])

# ---------------- Slide 6: Data Preprocessing ----------------
set_body(S[5].shapes[1].text_frame, [
    ("Feature engineering", 0),
    ("Store_Age_Years = 2025 − Store_Establishment_Year.", 1),
    ("Product_Id_char = first two characters of Product_Id (FD / DR / NC categories).", 1),
    ("Grouped 16 product types into two categories: Perishables vs Non-Perishables.", 1),
    ("Data cleaning", 0),
    ("Standardized inconsistent label: 'reg' → 'Regular' in Product_Sugar_Content.", 1),
    ("No missing values and no duplicate rows in the dataset.", 1),
    ("Data preparation for modeling", 0),
    ("Dropped identifiers/redundant columns: Product_Id, Product_Type, Store_Id, Store_Establishment_Year.", 1),
    ("Pipeline: OneHotEncoder(handle_unknown='ignore') on the categorical features; numeric columns are dropped (deployed model is categorical-only), fused with each model.", 1),
    ("70/30 train-test split (6,134 train / 2,629 test), random_state = 1.", 1),
])

# ---------------- Slide 7: Data Background and Contents ----------------
set_body(S[6].shapes[1].text_frame, [
    ("Dataset: 8,763 rows × 12 columns; no missing values, no duplicates.", 0),
    ("Product attributes: Product_Weight, Product_Sugar_Content, Product_Allocated_Area, Product_Type, Product_MRP.", 0),
    ("Store attributes: Store_Establishment_Year, Store_Size, Store_Location_City_Type, Store_Type.", 0),
    ("Identifiers: Product_Id, Store_Id (4 outlets: OUT001–OUT004).", 0),
    ("Target variable: Product_Store_Sales_Total (total sales of a product at a store).", 0),
    ("Coverage: 16 product types across 4 stores spanning 3 city tiers and 3 store sizes.", 0),
])

# ---------------- Slide 8: Model Building (+ base model table) ----------------
set_body(S[7].shapes[1].text_frame, [
    ("Built Random Forest and XGBoost regressors inside one preprocessing pipeline (one-hot encoded store/product categoricals); evaluated on the held-out test set.", 0),
    ("Both base models perform almost identically — Test R² ≈ 0.668, RMSE ≈ ₹616, MAPE ≈ 18.7% — a statistical tie, with Random Forest holding a slight, consistent edge.", 0),
    ("R² and RMSE are two views of the same errors (R² = 1 − RMSE²/Var(y)), so they rank the models the same; MAPE ≈ 18.7% agrees, confirming the tie.", 0),
    ("The models explain ~67% of sales variance, giving a solid baseline for demand planning.", 0),
])
add_table(S[7], [
    ["Model (base)", "Test R²", "RMSE (₹)", "MAE (₹)", "MAPE"],
    ["Random Forest", "0.668", "616", "485", "18.7%"],
    ["XGBoost", "0.668", "616", "485", "18.7%"],
], left=0.6, top=2.7, width=8.8, height=1.1,
   col_widths=[3.0, 1.45, 1.45, 1.45, 1.45], font_size=11, header_size=11)

# ---------------- Slide 9: Hyperparameter Tuning ----------------
set_body(S[8].shapes[1].text_frame, [
    ("Tuned both finalists with GridSearchCV (cv = 3, scoring = R²).", 0),
    ("Random Forest — best params: max_features = 'sqrt', n_estimators = 100, max_depth = None (the grid landed on the defaults).", 0),
    ("Test R² unchanged at 0.668; tuned Random Forest is identical to the base model.", 1),
    ("XGBoost — best params: gamma = 0, n_estimators = 100 (near-defaults).", 0),
    ("Test R² unchanged at 0.668 (RMSE ₹616) — no meaningful gain.", 1),
    ("Takeaway: tuning produced no measurable improvement — the models are already at the performance ceiling set by the available (categorical-only) features. Random Forest and XGBoost remain tied, and tuned Random Forest is taken forward as the selected model.", 0),
])

# ---------------- Slide 10: Model Performance Summary (+ leaderboard) ----------------
set_body(S[9].shapes[1].text_frame, [
    ("Final selected model: Tuned Random Forest (serialized with joblib and deployed).", 0),
    ("All four models are statistically tied — Test R² ≈ 0.668, differing only at the 4th decimal — but Random Forest holds a slight, consistent edge over XGBoost across every metric.", 0),
    ("Random Forest (tuned) was selected for deployment; the price-inclusive analysis (next slide) reinforces Random Forest as the stronger model family.", 0),
    ("Selected-model test performance: R² = 0.668, RMSE = ₹616, MAE = ₹485, MAPE = 18.7%.", 0),
])
add_table(S[9], [
    ["Model", "Test R²", "Adj R²", "RMSE (₹)", "MAE (₹)", "MAPE"],
    ["Random Forest (tuned)  ★ deployed", "0.668", "0.667", "616", "485", "18.7%"],
    ["Random Forest (base)", "0.668", "0.667", "616", "485", "18.7%"],
    ["XGBoost (base)", "0.668", "0.667", "616", "485", "18.7%"],
    ["XGBoost (tuned)", "0.668", "0.667", "616", "486", "18.7%"],
], left=0.5, top=2.4, width=9.0, height=1.8,
   col_widths=[2.85, 1.25, 1.25, 1.25, 1.2, 1.2], font_size=10.5, header_size=10.5)

# ---------------- Slide 11: Deployment ----------------
set_body(S[10].shapes[1].text_frame, [
    ("Tuned Random Forest serialized with joblib and served via a Flask REST API.", 0),
    ("Endpoint POST /v1/predict accepts product + store attributes and returns predicted sales as JSON.", 1),
    ("Streamlit web app provides a UI for single predictions and batch (CSV) scoring.", 0),
    ("Both the backend API and frontend are containerized with Docker and hosted on Hugging Face Spaces.", 0),
    ("Backend API: https://araujohn-rkt-superkart-backend.hf.space  (POST /v1/predict)", 0),
    ("Frontend Streamlit app: https://araujohn-rkt-superkart-app.hf.space", 0),
])

# ---------------- Slide 12: Appendix / pointers ----------------
S[11].shapes[0].text_frame.paragraphs[0].runs[0].text = "Key Takeaways & Business Recommendations"
set_body(S[11].shapes[1].text_frame, [
    ("Pricing & store format are the biggest sales levers in the data: Product_MRP correlates 0.79 with sales, and Departmental/Tier-1 stores sell ~3x Food Marts/Tier-3.", 0),
    ("Replicate winners: study OUT004 and Departmental/Tier-1 formats to lift underperforming outlets (OUT002, Food Marts, Tier-3).", 0),
    ("Model selection: Random Forest and XGBoost are statistically tied (Test R² ≈ 0.668), with a slight edge to Random Forest; tuned Random Forest was deployed.", 0),
    ("Highest-impact next step: add Product_MRP (price) to the model — a controlled test lifts Test R² from 0.67 to 0.93 and cuts typical error roughly in half (see Recommendation slide).", 0),
    ("Production-ready service (Flask API + Streamlit on Hugging Face); retrain periodically as new sales data arrives.", 0),
])

# ---------------- Slide 2: Agenda (append recommendation entry) ----------------
agenda_tf = None
for sh in S[1].shapes:
    if sh.has_text_frame and "Executive Summary" in sh.text_frame.text:
        agenda_tf = sh.text_frame
        break
if agenda_tf is not None:
    last = agenda_tf.paragraphs[-1]
    newp = agenda_tf.add_paragraph()
    newp.level = last.level
    nr = newp.add_run()
    nr.text = "Recommendation – Adding Product Price (MRP)"
    if last.runs:
        sr = last.runs[0]
        if sr.font.size:
            nr.font.size = sr.font.size
        if sr.font.bold is not None:
            nr.font.bold = sr.font.bold
        if sr.font.name:
            nr.font.name = sr.font.name
        try:
            nr.font.color.rgb = sr.font.color.rgb
        except Exception:
            pass

# ---------------- New slide: Recommendation - Adding Product Price (MRP) ----------------
rec = prs.slides.add_slide(prs.slide_layouts[2])  # TITLE_AND_BODY
rec.shapes.title.text = "Recommendation: Adding Product Price (MRP)"
rec_body = None
for ph in rec.placeholders:
    if ph.placeholder_format.idx == 1:
        rec_body = ph
        break
set_body(rec_body.text_frame, [
    ("Product_MRP (price) is the single strongest predictor of sales in EDA (correlation 0.79), yet the current production model excludes it per project scope.", 0),
    ("A controlled experiment adding Product_MRP improves every metric dramatically — and confirms Random Forest as the best model, widening its edge over XGBoost.", 0),
    ("Business impact: typical forecast error roughly halves (RMSE ₹616 → ₹291), enabling materially tighter inventory and supply-chain planning.", 0),
    ("Recommendation: incorporate Product_MRP (and Product_Weight) in the next model iteration once approved for use.", 0),
], base_size=13)
add_table(rec, [
    ["Feature set", "Best model", "Test R²", "RMSE (₹)", "MAPE"],
    ["Without price (current)", "Random Forest (tuned)", "0.668", "616", "18.7%"],
    ["With price (proposed)", "Random Forest (tuned)", "0.926", "291", "5.2%"],
], left=0.6, top=3.75, width=8.6, height=1.1,
   col_widths=[2.4, 2.2, 1.3, 1.4, 1.3], font_size=11, header_size=11)

# Position the new slide right after Deployment (becomes slide 12, before Key Takeaways)
xml_slides = prs.slides._sldIdLst
new_sldId = list(xml_slides)[-1]
xml_slides.remove(new_sldId)
xml_slides.insert(11, new_sldId)

prs.save(SRC)
print("Saved", SRC, "with", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
