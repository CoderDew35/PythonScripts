import pandas as pd
def fetch_and_analyze_data(months):
    results = []
    for month in months:
        query = f"SELECT * FROM dwh.telecom_churn_table WHERE ID_YEARMONTH = {month}"
        df = datasetFetch(cursor, sql_file=None, query=query)
       
        # Count occurrences of 0 and 1 for specified columns
        counts = {
            'T1ChurnFlag - 0': (df['T1ChurnFlag'] == 0).sum(),
            'T1ChurnFlag - 1': (df['T1ChurnFlag'] == 1).sum(),
            'MNPFlag - 0': (df['MNPFlag'] == 0).sum(),
            'MNPFlag - 1': (df['MNPFlag'] == 1).sum(),
            'DROP_FLAG - 0': (df['DROP_FLAG'] == 0).sum(),
            'DROP_FLAG - 1': (df['DROP_FLAG'] == 1).sum()
        }
       
        results.append((month, counts))
   
    # Create DataFrame for results
    result_df = pd.DataFrame(columns=['Month', 'Metric', 'Count'])
    for month, counts in results:
        for metric, count in counts.items():
            result_df = result_df.append({'Month': month, 'Metric': metric, 'Count': count}, ignore_index=True)
   
    # Save results to Excel file
    result_df.to_excel('count_results.xlsx', index=False)
   
    # Print results
    for month, counts in results:
        print(f"Results for month {month}:")
        for metric, count in counts.items():
            print(f"{metric}: {count}")
        print()
# Months to analyze
months = ['202401', '202312', '202311', '202310', '202309']
fetch_and_analyze_data(months)