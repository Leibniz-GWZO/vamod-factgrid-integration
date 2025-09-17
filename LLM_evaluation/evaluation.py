import pandas as pd
import numpy as np
from collections import defaultdict
import re

def clean_text(text):
    """Clean text by converting to lowercase and stripping whitespace"""
    if pd.isna(text) or text == '':
        return ''
    return str(text).lower().strip()

def extract_entities_from_row(row, entity_type):
    """Extract all entities of a specific type from a row"""
    entities = set()
    
    # Find all columns that match the pattern (e.g., "Genannte Person 1", "Genannte Person 2", etc.)
    pattern = rf"{entity_type} \d+"
    for col in row.index:
        if re.match(pattern, col):
            value = clean_text(row[col])
            if value and value != '':
                entities.add(value)
    
    return entities

def calculate_metrics(ground_truth_entities, predicted_entities):
    """Calculate precision, recall and F1-score"""
    if len(predicted_entities) == 0 and len(ground_truth_entities) == 0:
        return 1.0, 1.0, 1.0  # Perfect score if both are empty
    
    if len(predicted_entities) == 0:
        return 0.0, 0.0, 0.0  # No predictions made
    
    if len(ground_truth_entities) == 0:
        return 0.0, 0.0, 0.0  # No ground truth to match against
    
    # Calculate intersections
    true_positives = len(ground_truth_entities.intersection(predicted_entities))
    false_positives = len(predicted_entities - ground_truth_entities)
    false_negatives = len(ground_truth_entities - predicted_entities)
    
    # Calculate metrics
    precision = true_positives / (true_positives + false_positives) if (true_positives + false_positives) > 0 else 0.0
    recall = true_positives / (true_positives + false_negatives) if (true_positives + false_negatives) > 0 else 0.0
    f1_score = 2 * (precision * recall) / (precision + recall) if (precision + recall) > 0 else 0.0
    
    return precision, recall, f1_score

def evaluate_model(ground_truth_df, model_df, model_name):
    """Evaluate a single model against ground truth"""
    print(f"\n=== Evaluierung für {model_name} ===")
    
    # Entity types to evaluate
    entity_types = ["Genannte Person", "Empfängersitz", "Objektort"]
    
    overall_results = {}
    
    for entity_type in entity_types:
        print(f"\n--- {entity_type} ---")
        
        all_precisions = []
        all_recalls = []
        all_f1s = []
        
        # Match rows by "Nr. Rep"
        for nr_rep in ground_truth_df["Nr. Rep"].unique():
            if pd.isna(nr_rep):
                continue
                
            gt_row = ground_truth_df[ground_truth_df["Nr. Rep"] == nr_rep]
            model_row = model_df[model_df["Nr. Rep"] == nr_rep]
            
            if gt_row.empty or model_row.empty:
                continue
                
            # Extract entities for this row
            gt_entities = extract_entities_from_row(gt_row.iloc[0], entity_type)
            model_entities = extract_entities_from_row(model_row.iloc[0], entity_type)
            
            # Calculate metrics for this row
            precision, recall, f1 = calculate_metrics(gt_entities, model_entities)
            
            all_precisions.append(precision)
            all_recalls.append(recall)
            all_f1s.append(f1)
        
        # Calculate average metrics
        avg_precision = np.mean(all_precisions) if all_precisions else 0.0
        avg_recall = np.mean(all_recalls) if all_recalls else 0.0
        avg_f1 = np.mean(all_f1s) if all_f1s else 0.0
        
        overall_results[entity_type] = {
            'precision': avg_precision,
            'recall': avg_recall,
            'f1_score': avg_f1
        }
        
        print(f"  Precision: {avg_precision:.4f}")
        print(f"  Recall:    {avg_recall:.4f}")
        print(f"  F1-Score:  {avg_f1:.4f}")
    
    # Calculate macro-averaged metrics across all entity types
    macro_precision = np.mean([results['precision'] for results in overall_results.values()])
    macro_recall = np.mean([results['recall'] for results in overall_results.values()])
    macro_f1 = np.mean([results['f1_score'] for results in overall_results.values()])
    
    print(f"\n--- Macro-Averaged Metrics ---")
    print(f"  Precision: {macro_precision:.4f}")
    print(f"  Recall:    {macro_recall:.4f}")
    print(f"  F1-Score:  {macro_f1:.4f}")
    
    return overall_results, (macro_precision, macro_recall, macro_f1)

def main():
    # Load Excel file
    excel_file = "LLM_Evaluierung.xlsx"
    
    try:
        # Load all sheets
        print("Lade Excel-Datei...")
        ground_truth_df = pd.read_excel(excel_file, sheet_name="Ground_Truth")
        regex_df = pd.read_excel(excel_file, sheet_name="Regex")
        llama_df = pd.read_excel(excel_file, sheet_name="LLM llama-3.3")
        gpt4o_df = pd.read_excel(excel_file, sheet_name="LLM Gpt_4o")
        gpt41_df = pd.read_excel(excel_file, sheet_name="LLM Gpt_4.1")
        
        print(f"Ground Truth: {len(ground_truth_df)} Zeilen")
        print(f"Regex Modell: {len(regex_df)} Zeilen")
        print(f"Llama Modell: {len(llama_df)} Zeilen")
        print(f"GPT-4o Modell: {len(gpt4o_df)} Zeilen")
        print(f"GPT-4.1 Modell: {len(gpt41_df)} Zeilen")
        
        # Define models
        models = [
            ("Reguläre Ausdrücke + manuell annotierte Listen", regex_df),
            ("llama-3.3-70b-instruct", llama_df),
            ("gpt-4o", gpt4o_df),
            ("gpt-4.1", gpt41_df)
        ]
        
        # Store results for summary
        summary_results = {}
        
        # Evaluate each model
        for model_name, model_df in models:
            results, macro_metrics = evaluate_model(ground_truth_df, model_df, model_name)
            summary_results[model_name] = {
                'detailed': results,
                'macro': macro_metrics
            }
        
        # Print summary table
        print("\n" + "="*80)
        print("ZUSAMMENFASSUNG - MACRO-AVERAGED METRICS")
        print("="*80)
        print(f"{'Modell':<45} {'Precision':<12} {'Recall':<12} {'F1-Score':<12}")
        print("-"*80)
        
        for model_name in summary_results:
            macro_precision, macro_recall, macro_f1 = summary_results[model_name]['macro']
            print(f"{model_name:<45} {macro_precision:<12.4f} {macro_recall:<12.4f} {macro_f1:<12.4f}")
        
        # Print detailed table by entity type
        print("\n" + "="*80)
        print("DETAILLIERTE ERGEBNISSE NACH ENTITY-TYP")
        print("="*80)
        
        entity_types = ["Genannte Person", "Empfängersitz", "Objektort"]
        
        for entity_type in entity_types:
            print(f"\n--- {entity_type} ---")
            print(f"{'Modell':<45} {'Precision':<12} {'Recall':<12} {'F1-Score':<12}")
            print("-"*80)
            
            for model_name in summary_results:
                results = summary_results[model_name]['detailed'][entity_type]
                precision = results['precision']
                recall = results['recall']
                f1_score = results['f1_score']
                print(f"{model_name:<45} {precision:<12.4f} {recall:<12.4f} {f1_score:<12.4f}")
        
    except FileNotFoundError:
        print(f"Fehler: Die Datei '{excel_file}' wurde nicht gefunden.")
        print("Stellen Sie sicher, dass sich die Datei im gleichen Verzeichnis wie das Skript befindet.")
    except Exception as e:
        print(f"Fehler beim Laden der Datei: {e}")

if __name__ == "__main__":
    main()