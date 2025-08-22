# Question Label Processor

This project processes raw question labels from a Word document, predicting and turning them into standardized canonical labels using a trained SVM model, and compares a separate list of ground truth labels against the raw labels to determine whether each is present.

The output is two Word documents:
- ProcessedLabels.docx – Contains raw labels alongside their predicted canonical labels.

- GroundTruth.docx – Contains ground truth labels with a "Found/Not Found" status depending on whether they appear in the raw labels.

## Features

- Converts raw, unstructured labels into a standardized canonical format.
- Uses a character-level TF-IDF vectorizer and SVM classifier for flexible recognition of label variations.
- Maintains a separate ground truth list that automatically updates with "Found/Not Found" results.
- Works with Microsoft Word .docx files for both input and output.

## Requirements

- Python 3.8+
- Install dependencies with:
  - pip install scikit-learn pandas python-docx openpyxl

## File Structure

- raw_labels.docx: Input Word document containing raw labels, one per line.

- GroundTruth.docx: Input/output Word document containing ground truth labels.

  - On first run, this may contain a simple header ("Ground Truth") followed by labels.

  - The script will convert this into a table and maintain it going forward.

- question_label_variations_expanded.xlsx: Training dataset containing raw_label and canonical_label columns.

- ProcessedLabels.docx: Output Word document with raw labels and predicted canonical labels.

## Usage

Prepare the training dataset (question_label_variations_expanded.xlsx) with at least two columns:

- raw_label (ex: "Q1", "question one")

- canonical_label (ex: "1")

Place your raw labels into raw_labels.docx, one per line.

Prepare GroundTruth.docx with a table, the first column (with a header of "Ground Truth") having all the base labels, and a blank second column (with a header of "Found/Not Found")

Run the script. It will:

- Train the SVM model on your Excel dataset.

- Generate ProcessedLabels.docx with predictions.

- Update GroundTruth.docx into a table, filling the "Found/Not Found" column.

## Adding New Ground Truths

- After the first run, GroundTruth.docx will be a table.

- To add more labels, insert new rows into the Ground Truth column of the table.

- Rerun the script to update statuses.
