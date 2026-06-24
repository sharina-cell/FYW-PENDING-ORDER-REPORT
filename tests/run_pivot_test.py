import pandas as pd
from report_engine import build_pivot


def main():
    data = [
        {'Nickname': 'Alpha', 'Brand': 'MELISSA', 'Order ID': 'O1', 'FYW SLA': '24/06/2026'},
        {'Nickname': 'Alpha', 'Brand': 'MELISSA', 'Order ID': 'O2', 'FYW SLA': '22/06/2026'},
        {'Nickname': 'Beta',  'Brand': 'IPANEMA', 'Order ID': 'O3', 'FYW SLA': '23/06/2026'},
        {'Nickname': 'Beta',  'Brand': 'IPANEMA', 'Order ID': 'O4', 'FYW SLA': '22/06/2026'},
    ]

    df = pd.DataFrame(data)
    pivot, all_cols = build_pivot(df)

    print('all_cols:', all_cols)
    print('pivot.columns:', list(pivot.columns))
    print('\nPivot frame:')
    print(pivot)


if __name__ == '__main__':
    main()
