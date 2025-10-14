import pandas as pd

def merge_files(file1, file2, on_col, output_path):
    """
    Συγχώνευση δύο Excel αρχείων με βάση μια κοινή στήλη.

    :param file1: πρώτο Excel αρχείο
    :param file2: δεύτερο Excel αρχείο
    :param on_col: στήλη για merge
    :param output_path: που θα αποθηκευτεί το αποτέλεσμα (.xlsx ή .csv)
    """
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    if on_col not in df1.columns:
        raise KeyError(f"Η στήλη '{on_col}' δεν βρέθηκε στο πρώτο αρχείο.")
    if on_col not in df2.columns:
        raise KeyError(f"Η στήλη '{on_col}' δεν βρέθηκε στο δεύτερο αρχείο.")

    merged = pd.merge(df1, df2, on=on_col, how="inner")

    # Αποθήκευση ανάλογα με το extension
    if output_path.endswith(".csv"):
        merged.to_csv(output_path, index=False)
    else:
        merged.to_excel(output_path, index=False)

    return output_path


