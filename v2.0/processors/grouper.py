import pandas as pd

def group_file(file_path, group_col, value_col, agg_func, output_path):
    """
    Ομαδοποίηση αρχείου Excel με διαφορετικές συναρτήσεις συγκέντρωσης.

    :param file_path: input Excel αρχείο
    :param group_col: στήλη για group by
    :param value_col: στήλη με αριθμητικές τιμές
    :param agg_func: 'sum', 'mean', 'count', 'max', 'min'
    :param output_path: που θα αποθηκευτεί το αποτέλεσμα (.xlsx ή .csv)
    """
    df = pd.read_excel(file_path)

    if group_col not in df.columns:
        raise KeyError(f"Η στήλη '{group_col}' δεν βρέθηκε.")
    if value_col not in df.columns:
        raise KeyError(f"Η στήλη '{value_col}' δεν βρέθηκε.")

    if agg_func not in ["sum", "mean", "count", "max", "min"]:
        raise ValueError(f"Μη υποστηριζόμενη συνάρτηση: {agg_func}")

    grouped = df.groupby(group_col)[value_col].agg(agg_func).reset_index()

    # Αποθήκευση ανάλογα με το extension
    if output_path.endswith(".csv"):
        grouped.to_csv(output_path, index=False)
    else:
        grouped.to_excel(output_path, index=False)

    return output_path
