#!/usr/bin/python3

import sys
import os
import pandas as pd


def convert(file):
    if os.path.exists(os.path.dirname(file.path) + '/gis/' +
                      file.name + ' ΓΙΑ gis.xlsx'):
        print("Already converted. Skipping...")
        return

    # Διάβασε το αρχείο μετατρέποντας τα πάντα σε str
    # και αγνοώντας τα κενά κελιά
    df = pd.read_excel(file.path, dtype=str, keep_default_na=False)
    for i in df.columns.tolist():
        df[i] = df[i].str.strip()
    df.rename(columns={'ΑΡΙΘΜΟΣ ΜΗΤΡΩΟΥ ΜΑΘΗΤΗ': 'Αρ. μητρώου',
                       'ΔΗΜΟΣ': 'Διεύθυνση, περιοχή',
                       'Τ.Κ.': 'Διεύθυνση, Τ.Κ.'},
              inplace=True)
    df['Επώνυμο μαθητή'] = ""
    df['Όνομα μαθητή'] = ""
    df['Όνομα πατέρα'] = ""
    df['Διεύθυνση, οδός - αριθμός'] = df['Δ/ΝΣΗ ΚΑΤΟΙΚΙΑΣ ΜΑΘΗΤΗ'].str.cat(
        df['ΑΡΙΘΜΟΣ'], sep=' ', na_rep='')
    result = map(lambda x: 'ΑΔΕΡΦΟΣ/Η ΣΤΟ ' + x if x != '' else '',
                 df['ΓΥΜΝΑΣΙΟ ΑΔΕΛΦΟΥ/ΗΣ'].tolist())
    df['ΓΥΜΝΑΣΙΟ ΑΔΕΛΦΟΥ/ΗΣ'] = pd.Series(list(result))
    result = map(lambda x: 'ΤΜΗΜΑ ΕΝΤΑΞΗΣ ' + x if x != '' else '',
                 df['ΦΟΙΤΗΣΗ ΣΕ ΤΜΗΜΑ ΕΝΤΑΞΗΣ'].tolist())
    df['ΦΟΙΤΗΣΗ ΣΕ ΤΜΗΜΑ ΕΝΤΑΞΗΣ'] = pd.Series(list(result))
    df['Πληροφορίες'] = pd.Series(['\n'.join(i)
                                  for i in zip(df['ΓΥΜΝΑΣΙΟ ΑΔΕΛΦΟΥ/ΗΣ'],
                                               df['ΦΟΙΤΗΣΗ ΣΕ ΤΜΗΜΑ ΕΝΤΑΞΗΣ'],
                                               df['ΠΑΡΑΤΗΡΗΣΕΙΣ'])])
    df['Πληροφορίες'] = df['Πληροφορίες'].str.strip()
    df['Σχολείο Τοποθέτησης'] = ""
    df['email Γονέα'] = ""
    df['Διεύθυνση για Google (προαιρετικό)'] = ""
    df['Συντεταγμένες Διεύθυνσης (προαιρετικό)'] = ""
    newdf = df[["Αρ. μητρώου", "Επώνυμο μαθητή", "Όνομα μαθητή",
                "Όνομα πατέρα", "Διεύθυνση, οδός - αριθμός", "Διεύθυνση, Τ.Κ.",
                "Διεύθυνση, περιοχή", "Πληροφορίες", "Σχολείο Τοποθέτησης",
                "email Γονέα", "Διεύθυνση για Google (προαιρετικό)",
                "Συντεταγμένες Διεύθυνσης (προαιρετικό)"]]
    try:
        os.mkdir(os.path.dirname(file.path)+"/gis")
    except FileExistsError:
        pass
    with pd.ExcelWriter(os.path.dirname(file.path) + '/gis/' +
                        file.name + ' ΓΙΑ gis.xlsx') as writer:
        newdf.to_excel(writer)
        writer.save()


def scandir(dirlist):
    while dirlist:
        with os.scandir(dirlist.pop()) as it:
            for entry in it:
                if entry.is_dir() and entry.name != 'gis':
                    print("Found dir", entry.name)
                    dirlist.append(entry.path)
                else:
                    filename = entry.name.lower()
                    if filename.endswith('.xls') or \
                       filename.endswith('.xlsx'):
                        print("Converting file", entry.name)
                        convert(entry)


workdir = sys.argv[1]
scandir([workdir])
