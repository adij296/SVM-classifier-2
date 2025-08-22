from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC
from sklearn.feature_extraction.text import TfidfVectorizer
import pandas as pd
import os
from openpyxl import load_workbook
from docx import Document

# Load data from spreadsheet
EXCEL = 'question_label_variations_expanded.xlsx'
data = pd.read_excel(EXCEL)

# Assign labels
X = data['raw_label'].astype(str)         # e.g., "Q1", "question one"
y = data['canonical_label'].astype(str)   # ensure labels are strings

# Train-test split
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)

#Vectorization
vectorizer = TfidfVectorizer(analyzer='char_wb', ngram_range=(2, 4))
X_train_tfidf = vectorizer.fit_transform(X_train)
X_test_tfidf = vectorizer.transform(X_test)

# Train SVM
model = SVC(kernel='rbf', C=1.0, gamma='scale')
model.fit(X_train_tfidf, y_train)

# Load in input labels
INPUTDOC = 'raw_labels.docx'
raw_doc = Document(INPUTDOC)
raw_labels = [p.text.strip() for p in raw_doc.paragraphs if p.text.strip()]
raw_df = pd.DataFrame({"raw_label": raw_labels})

raw_df["predicted_canonical"] = model.predict(vectorizer.transform(raw_df["raw_label"]))

#Save ProcessedLabels
processed_doc = Document()
table = processed_doc.add_table(rows=1, cols=2)
hdr = table.rows[0].cells
hdr[0].text = "Raw Label"
hdr[1].text = "Predicted Canonical Label"

for _, row in raw_df.iterrows():
    cells = table.add_row().cells
    cells[0].text = str(row["raw_label"])
    cells[1].text = str(row["predicted_canonical"])

PROCESSEDDOC = 'ProcessedLabels.docx'
processed_doc.save(PROCESSEDDOC)
print("Processed Labels saved")

GROUNDTRUTHDOC = 'GroundTruth.docx'
raw_labels_set = set(raw_df["raw_label"].astype(str))

# Load existing GroundTruth doc
gt_doc = Document(GROUNDTRUTHDOC)

# If a table exists, use it, otherwise create a new table with headers
if gt_doc.tables:
    table = gt_doc.tables[0]
else:
    table = gt_doc.add_table(rows=1, cols=2)
    table.rows[0].cells[0].text = "Ground Truth"
    table.rows[0].cells[1].text = "Found/Not Found"

# Collect existing ground truth labels in the table
existing_labels = set()
for row in table.rows[1:]:  # skip header
    gt_label = row.cells[0].text.strip()
    existing_labels.add(gt_label)

# Gather all labels to process: existing + any new ones from elsewhere
for row in table.rows[1:]:
    gt_label = row.cells[0].text.strip()
    row.cells[1].text = "Found" if gt_label in raw_labels_set else "Not Found"

gt_doc.save(GROUNDTRUTHDOC)
print("Ground Truth updated with Found/Not Found")
