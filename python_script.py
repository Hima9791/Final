import pandas as pd
from helper import load_file, similarity_ratio, normalize_series

# GitHub raw URLs
MASTER_URL = "https://raw.githubusercontent.com/USERNAME/REPO/BRANCH/MasterSeriesHistory.xlsx"
RULES_URL = "https://raw.githubusercontent.com/USERNAME/REPO/BRANCH/SampleSeriesRules.xlsx"

def compare_requested_series_from_github(comparison_path, top_n=2):
    df_master = load_file(MASTER_URL)
    df_rules = load_file(RULES_URL)
    df_comparison = load_file(comparison_path)

    results = []
    for _, row in df_comparison.iterrows():
        comparison_series = normalize_series(row["Series"])
        matches = []
        for _, master_row in df_master.iterrows():
            master_series = normalize_series(master_row["Series"])
            ratio = similarity_ratio(comparison_series, master_series)
            matches.append((master_series, ratio))
        matches.sort(key=lambda x: x[1], reverse=True)
        top_matches = matches[:top_n]
        results.append({"Requested": comparison_series, "Matches": top_matches})

    return pd.DataFrame(results)

# Other update/delete functions can stay here...

def update_master_series(update_path, master_path):
    df_update = load_file(update_path)
    df_master = load_file(master_path)
    df_master = pd.concat([df_master, df_update]).drop_duplicates().reset_index(drop=True)
    df_master.to_excel(master_path, index=False)

def delete_from_master_series(delete_path, master_path):
    df_delete = load_file(delete_path)
    df_master = load_file(master_path)
    df_master = df_master[~df_master["Series"].isin(df_delete["Series"])]
    df_master.to_excel(master_path, index=False)

def update_series_rules(update_path, rules_path):
    df_update = load_file(update_path)
    df_rules = load_file(rules_path)
    df_rules = pd.concat([df_rules, df_update]).drop_duplicates().reset_index(drop=True)
    df_rules.to_excel(rules_path, index=False)

def delete_from_series_rules(delete_path, rules_path):
    df_delete = load_file(delete_path)
    df_rules = load_file(rules_path)
    df_rules = df_rules[~df_rules["Series"].isin(df_delete["Series"])]
    df_rules.to_excel(rules_path, index=False)
