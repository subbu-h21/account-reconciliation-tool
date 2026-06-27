import pandas as pd
import numpy as np
import re

# 1st: receivements discrepancy
def process_receivements(ac_path, books_path, timestamp):
    # engines
    ac_eng = 'openpyxl' if ac_path.endswith('.xlsx') else 'xlrd'
    ob_eng = 'openpyxl' if books_path.endswith('.xlsx') else 'xlrd'
    # load account statement
    ac = (pd.read_excel(ac_path, header=None, engine=ac_eng)
            .drop(index=range(14)).dropna(axis=1, how='all').reset_index(drop=True))
    ac.columns = ac.iloc[0]; ac = ac[1:].reset_index(drop=True).drop('Cheque No', axis=1)
    ac.columns = ['Date','Particular','Given','Received','Balance']
    # load our books
    ob = (pd.read_excel(books_path, header=None, engine=ob_eng)
            .dropna(axis=1, how='all')
            .pipe(lambda df: df.rename(columns=df.iloc[0])).drop(index=0)
            .pipe(lambda df: df[~df['Particular'].str.contains('Opening balance|Closing balance', case=False, na=False)])
            .reset_index(drop=True))
    # filter receivements
    recv_ac = ac[ac['Received'].notna()].copy()
    recv_ob = ob[ob['Debit'].notna() & (ob['Debit'] != 0)].copy(); recv_ob['Particular'] = recv_ob['Particular'].str.strip()
    # helpers
    def trim(s): return s.strip()
    def inside(text): m = re.search(r'\((.*?)\)?$', str(text)); return m.group(1) if m else str(text)
    recv_ac['SupplierName'] = recv_ac['Particular'].apply(trim)
    recv_ob['Particular']    = recv_ob['Particular'].apply(inside)
    # drop UPI entries at source
    recv_ac = recv_ac[~recv_ac['SupplierName'].str.contains('UPI', na=False)]
    recv_ob = recv_ob[~recv_ob['Particular'].str.contains('UPI', na=False)]
    # parse dates
    recv_ac['Date'] = pd.to_datetime(recv_ac['Date'], errors='coerce', dayfirst=True).dt.date
    recv_ob['Date'] = pd.to_datetime(recv_ob['Date'], errors='coerce', dayfirst=True).dt.date
    # parse amounts
    recv_ac['Received'] = recv_ac['Received'].replace({'\\$':'','\\,':'','\\s+':''}, regex=True).pipe(pd.to_numeric, errors='coerce')
    recv_ob['Debit']    = recv_ob['Debit'].replace({'\\$':'','\\,':'','\\s+':''}, regex=True).pipe(pd.to_numeric, errors='coerce')
    # collect all individual entries per date
    rows, prev = [], None
    for d in sorted(set(recv_ac['Date'].dropna()) | set(recv_ob['Date'].dropna())):
        a = recv_ac[recv_ac['Date']==d]; b = recv_ob[recv_ob['Date']==d]
        only_ac = list(zip(a['SupplierName'], a['Received']))
        only_ob = list(zip(b['Particular'],   b['Debit']))
        L = max(len(only_ac), len(only_ob), 1)
        only_ac += [(None, None)] * (L - len(only_ac))
        only_ob += [(None, None)] * (L - len(only_ob))
        if prev is not None and d != prev:
            rows.append({'Date': None, 'Books': None, 'Debit': None, 'Received': None, 'Bank': None})
        for (acn, acamt), (obn, obamt) in zip(only_ac, only_ob):
            rows.append({'Date': d, 'Books': obn, 'Debit': obamt, 'Received': acamt, 'Bank': acn})
        prev = d
    return pd.DataFrame(rows)[['Date', 'Books', 'Debit', 'Received', 'Bank']]

# 2nd: payments discrepancy
def process_payments(ac_path, books_path, timestamp):
    ac_eng = 'openpyxl' if ac_path.endswith('.xlsx') else 'xlrd'
    ob_eng = 'openpyxl' if books_path.endswith('.xlsx') else 'xlrd'
    ac = (pd.read_excel(ac_path, header=None, engine=ac_eng).drop(index=range(14)).dropna(axis=1, how='all').reset_index(drop=True))
    ac.columns = ac.iloc[0]; ac = ac[1:].reset_index(drop=True).drop('Cheque No', axis=1)
    ac.columns = ['Date','Particular','Given','Received','Balance']
    ob = (pd.read_excel(books_path, header=None, engine=ob_eng).dropna(axis=1, how='all')
            .pipe(lambda df: df.rename(columns=df.iloc[0])).drop(index=0)
            .pipe(lambda df: df[~df['Particular'].str.contains('Opening balance|Closing balance', case=False, na=False)]).reset_index(drop=True))
    pay_ac = ac[ac['Given'].notna()].copy(); pay_ob = ob[ob['Credit'].notna() & (ob['Credit'] != 0)].copy()
    pay_ob['Particular'] = pay_ob['Particular'].str.strip()
    # drop UPI entries at source
    pay_ac = pay_ac[~pay_ac['Particular'].str.contains('UPI', na=False)]
    pay_ob = pay_ob[~pay_ob['Particular'].str.contains('UPI', na=False)]
    def name_clean(desc): return re.sub(r"(NEFT-|MClick/To\s+)", "", str(desc)).strip()
    def sup_name(txn): parts=str(txn).split('/'); return parts[0].strip() if len(parts)>1 else str(txn)
    pay_ac['Particular']=pay_ac['Particular'].apply(name_clean); pay_ac['SupplierName']=pay_ac['Particular'].apply(sup_name)
    pay_ac['Date']=pd.to_datetime(pay_ac['Date'],errors='coerce',dayfirst=True).dt.date
    pay_ob['Date']=pd.to_datetime(pay_ob['Date'],errors='coerce',dayfirst=True).dt.date
    pay_ac['Given']=pay_ac['Given'].replace({'\\$':'','\\,':'','\\s+':''}, regex=True).pipe(pd.to_numeric, errors='coerce')
    pay_ob['Credit']=pay_ob['Credit'].replace({'\\$':'','\\,':'','\\s+':''}, regex=True).pipe(pd.to_numeric, errors='coerce')
    rows,prev=[],None
    for d in sorted(set(pay_ac['Date'].dropna()) | set(pay_ob['Date'].dropna())):
        a=pay_ac[pay_ac['Date']==d]; b=pay_ob[pay_ob['Date']==d]
        only_ac=list(zip(a['SupplierName'], a['Given']))
        only_ob=list(zip(b['Particular'],   b['Credit']))
        L=max(len(only_ac),len(only_ob),1)
        only_ac+=[(None,None)]*(L-len(only_ac)); only_ob+=[(None,None)]*(L-len(only_ob))
        if prev is not None and d!=prev:
            rows.append({'Date':None,'Books':None,'Credit':None,'Given':None,'Bank':None})
        for (acn,acamt),(obn,obamt) in zip(only_ac,only_ob):
            rows.append({'Date':d,'Books':obn,'Credit':obamt,'Given':acamt,'Bank':acn})
        prev=d
    final=pd.DataFrame(rows)[['Date','Books','Credit','Given','Bank']]
    return final

def process_summary(ac_path, books_path):
    # Determine engines
    ac_eng = 'openpyxl' if ac_path.endswith('.xlsx') else 'xlrd'
    ob_eng = 'openpyxl' if books_path.endswith('.xlsx') else 'xlrd'

    # Load and normalize account statement
    ac = (pd.read_excel(ac_path, header=None, engine=ac_eng)
            .drop(index=range(14))
            .dropna(axis=1, how='all')
            .reset_index(drop=True))
    ac.columns = ac.iloc[0]
    ac = ac[1:].reset_index(drop=True).drop('Cheque No', axis=1)
    ac.columns = ['Date','Particular','Given','Received','Balance']
    ac['Date'] = pd.to_datetime(ac['Date'], errors='coerce', dayfirst=True)

    # Load and normalize our books
    ob = (pd.read_excel(books_path, header=None, engine=ob_eng)
            .dropna(axis=1, how='all')
            .pipe(lambda df: df.rename(columns=df.iloc[0])).drop(index=0)
            .pipe(lambda df: df[~df['Particular'].str.contains('Opening balance|Closing balance', case=False, na=False)])
            .reset_index(drop=True))
    ob['Date'] = pd.to_datetime(ob['Date'], errors='coerce', dayfirst=True)

    # Clean numeric columns
    ac[['Given','Received']] = (ac[['Given','Received']]
        .replace({'\\$':'','\\,':'','\\s+':''}, regex=True)
        .replace([None,'0'], np.nan)
        .apply(pd.to_numeric, errors='coerce'))
    ob[['Credit','Debit']] = (ob[['Credit','Debit']]
        .replace({'\\$':'','\\,':'','\\s+':''}, regex=True)
        .replace([None,0], np.nan)
        .apply(pd.to_numeric, errors='coerce'))

    ac['Balance'] = (
        ac['Balance']
        .str.replace(',', '', regex=True)
        .astype(float)
        .abs()
    )
    ob['Balance'] = (
        ob['Balance']
        .str.replace('Cr', '', regex=False)
        .str.replace(',', '', regex=False)
        .str.replace(r'\s+', '', regex=True)
    )
    ob['Balance'] = pd.to_numeric(ob['Balance'], errors='coerce').abs()

    # Gather all unique dates
    all_dates = sorted(set(ac['Date'].dropna().dt.date.unique()) |
                       set(ob['Date'].dropna().dt.date.unique()))

    rows = []
    for d in all_dates:
        ac_d = ac[ac['Date'].dt.date == d]
        ob_d = ob[ob['Date'].dt.date == d]

        rows.append(((d, 'total count'), {
            'account': len(ac_d),
            'our_book': len(ob_d),
        }))
        rows.append(((d, 'debit/received entries'), {
            'account': ac_d['Received'].count(),
            'our_book': ob_d['Debit'].count(),
        }))
        rows.append(((d, 'credit/given entries'), {
            'account': ac_d['Given'].count(),
            'our_book': ob_d['Credit'].count(),
        }))
        rows.append(((d, 'debit/received total'), {
            'account': ac_d['Received'].sum(),
            'our_book': ob_d['Debit'].sum(),
        }))
        rows.append(((d, 'credit/given total'), {
            'account': ac_d['Given'].sum(),
            'our_book': ob_d['Credit'].sum(),
        }))
        rows.append(((d, 'Closing Balance'), {
            'account': ac_d['Balance'].iloc[0] if not ac_d.empty else None,
            'our_book': ob_d['Balance'].iloc[-1] if not ob_d.empty else None,
        }))

    # Build summary DataFrame
    idx = pd.MultiIndex.from_tuples([r[0] for r in rows], names=['Date','Metric'])
    df = pd.DataFrame([r[1] for r in rows], index=idx)
    df['Difference'] = df['account'] - df['our_book']
    # df = df.reset_index()
    return df
