import streamlit as st
import pandas as pd
import datetime
import re
from io import BytesIO
from pandas import ExcelWriter
import uuid

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")
st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    if 'TIME' not in df.columns:
        return pd.DataFrame()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df = df[df['DATE'].dt.weekday != 6]  # Exclude Sundays
    return df

def to_excel(summary_groups):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'}),
            'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True}),
            'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'}),
            'percent': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00%'}),
            'date': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'}),
            'time': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'hh:mm:ss'})
        }

        for sheet_name, df_dict in summary_groups.items():
            worksheet = workbook.add_worksheet(str(sheet_name))
            current_row = 0

            for title, df in df_dict.items():
                if df.empty:
                    continue
                df_for_excel = df.copy()
                for col in ['RPC RATE', 'PTP RATE', 'CVR']:
                    if col in df_for_excel.columns:
                        df_for_excel[col] = df_for_excel[col].str.rstrip('%').astype(float) / 100 if df_for_excel[col].dtype == 'object' else df_for_excel[col]

                worksheet.merge_range(current_row, 0, current_row, len(df.columns) - 1, title, formats['title'])
                current_row += 1

                for col_num, col_name in enumerate(df_for_excel.columns):
                    worksheet.write(current_row, col_num, col_name, formats['header'])
                    max_len = max(df_for_excel[col_name].astype(str).str.len().max(), len(col_name)) + 2
                    worksheet.set_column(col_num, col_num, max_len)

                current_row += 1

                for row_num in range(len(df_for_excel)):
                    for col_num, col_name in enumerate(df_for_excel.columns):
                        value = df_for_excel.iloc[row_num, col_num]
                        if col_name == 'PERIOD':
                            worksheet.write(current_row + row_num, col_num, value, formats['date'])
                        elif col_name in ['WORKED ACCOUNT', 'TOTAL CONNECTED', 'RPC', 'PTP', 'PTP REMINDER', 'PAYMENT REMINDER', 'CONFIRMED', 'PTP AMOUNT', 'CONFIRMED AMOUNT']:
                            worksheet.write(current_row + row_num, col_num, value, formats['comma'])
                        elif col_name in ['TOTAL TALK TIME', 'RPC TALK TIME', 'PTP TALK TIME']:
                            worksheet.write(current_row + row_num, col_num, value, formats['time'])
                        elif col_name in ['RPC RATE', 'PTP RATE', 'CVR']:
                            worksheet.write(current_row + row_num, col_num, value, formats['percent'])
                        else:
                            worksheet.write(current_row + row_num, col_num, value, formats['center'])

                current_row += len(df_for_excel) + 2

    return output.getvalue()

def process_file(df):
    if df.empty:
        return df
    df = df[df['REMARK BY'] != 'SPMADRID']
    df = df[~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
    df = df[~df['STATUS'].str.contains('ABORT', na=False)]
    df = df[~df['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True)]
    df = df[~df['REMARK'].str.contains('Broadcast', case=False, na=False)]
    excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                        "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -", "File Unhold"]
    escaped_remarks = [re.escape(remark) for remark in excluded_remarks]
    df = df[~df['REMARK'].str.contains('|'.join(escaped_remarks), case=False, na=False)]
    df['CARD NO.'] = df['CARD NO.'].astype(str)
    return df

def format_seconds_to_hms(seconds):
    seconds = int(seconds)
    hours, minutes = seconds // 3600, (seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}:{seconds % 60:02d}"

def calculate_agent_productivity_summary(df):
    summary_columns = ['PERIOD', 'AGENT', 'CAMPAIGN', 'WORKED ACCOUNT', 'TOTAL CONNECTED', 'TOTAL TALK TIME', 
                      'RPC', 'RPC TALK TIME', 'PTP', 'PTP TALK TIME', 'PTP REMINDER', 'PAYMENT REMINDER', 
                      'CONFIRMED', 'RPC RATE', 'PTP RATE', 'CVR', 'PTP AMOUNT', 'CONFIRMED AMOUNT']
    summary_table = pd.DataFrame(columns=summary_columns)
    
    if df.empty:
        return summary_table

    df_filtered = df[~df['REMARK BY'].str.upper().eq('SYSTEM')].copy()
    df_filtered['DATE'] = pd.to_datetime(df_filtered['DATE'], errors='coerce').dt.date
    df_filtered = df_filtered[df_filtered['DATE'].notna()]
    
    if df_filtered.empty:
        return summary_table

    # Explicitly set PTP DATE column
    ptp_date_col = 'PTP DATE' if 'PTP DATE' in df_filtered.columns else None
    if not ptp_date_col:
        st.warning("PTP DATE column not found in the data. PTP REMINDER and PAYMENT REMINDER will be set to 0.")
        ptp_date_col = None

    # Get PTP accounts
    ptp_excluded_statuses = ['PTP FF UP', 'PTP FOLLOW UP', 'PTP REMINDER', 'PTP FF UP', 'PTP FOLLOW UP', 
                            'PYMT REMINDER (FF UP)', 'PAYMENT REMINDER (FF UP)', 'PTP_FF UP']
    ptp_df = df_filtered[(df_filtered['STATUS'].str.contains('PTP', case=False, na=False, regex=False)) & 
                        (~df_filtered['STATUS'].str.contains('|'.join(ptp_excluded_statuses), case=False, na=False, regex=False)) & 
                        (pd.to_numeric(df_filtered['PTP AMOUNT'], errors='coerce') > 0)]
    # Remove deduplication to count all PTP occurrences
    # ptp_df = ptp_df.drop_duplicates(subset=['DATE', 'ACCOUNT NO.', 'PTP AMOUNT'])
    if ptp_date_col:
        ptp_df[ptp_date_col] = pd.to_datetime(ptp_df[ptp_date_col], errors='coerce')

    summary_rows = []
    for (agent, campaign), group in df_filtered.groupby(['REMARK BY', 'CLIENT']):
        min_date = group['DATE'].min()
        max_date = group['DATE'].max()
        period = f"{min_date.strftime('%b %d %Y')} - {max_date.strftime('%b %d %Y')}" if min_date and max_date else ""

        # Non-unique counts for WORKED ACCOUNT, TOTAL CONNECTED, and RPC
        worked_account = len(group[group['REMARK TYPE'] == 'Outgoing'])
        total_connected_df = group[pd.to_numeric(group['TALK TIME DURATION'], errors='coerce') > 0]
        total_connected = len(total_connected_df)
        total_talk_seconds = pd.to_numeric(total_connected_df['TALK TIME DURATION'], errors='coerce').sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds) if total_talk_seconds > 0 else "00:00:00"

        rpc_df = group[group['STATUS'].str.contains('POS CLIENT -|POSITIVE CLIENT -', case=False, na=False, regex=True)]
        rpc = len(rpc_df)
        rpc_talk_seconds = pd.to_numeric(rpc_df['TALK TIME DURATION'], errors='coerce').sum()
        rpc_talk_time = format_seconds_to_hms(rpc_talk_seconds) if rpc_talk_seconds > 0 else "00:00:00"

        agent_ptp_df = ptp_df[(ptp_df['REMARK BY'] == agent) & (ptp_df['CLIENT'] == campaign)]
        # Non-unique count for PTP
        ptp = len(agent_ptp_df)
        ptp_talk_seconds = pd.to_numeric(agent_ptp_df['TALK TIME DURATION'], errors='coerce').sum()
        ptp_talk_time = format_seconds_to_hms(ptp_talk_seconds) if ptp_talk_seconds > 0 else "00:00:00"
        ptp_amount = pd.to_numeric(agent_ptp_df['PTP AMOUNT'], errors='coerce').sum()

        # Calculate PTP REMINDER and PAYMENT REMINDER
        ptp_reminder = 0
        payment_reminder = 0
        if ptp_date_col:
            ptp_reminder_statuses = ['PTP_FF UP', 'PYMT REMINDER (FF UP)', 'PAYMENT REMINDER (FF UP)']
            reminder_df = df_filtered[(df_filtered['REMARK BY'] == agent) & 
                                    (df_filtered['CLIENT'] == campaign) & 
                                    (df_filtered['STATUS'].str.contains('|'.join(ptp_reminder_statuses), case=False, na=False, regex=False))]
            reminder_df['DATE'] = pd.to_datetime(reminder_df['DATE'], errors='coerce')
            reminder_df[ptp_date_col] = pd.to_datetime(reminder_df[ptp_date_col], errors='coerce')

            # Get PTP accounts for this agent and campaign
            ptp_accounts = agent_ptp_df[['ACCOUNT NO.', ptp_date_col]].drop_duplicates()
            # Rename PTP DATE to avoid merge conflicts
            ptp_accounts = ptp_accounts.rename(columns={ptp_date_col: 'PTP_DATE_ORIG'})
            
            # PTP REMINDER: Count reminders before PTP DATE for PTP accounts
            ptp_reminder_df = reminder_df.merge(ptp_accounts, on='ACCOUNT NO.', how='inner')
            ptp_reminder_df = ptp_reminder_df[ptp_reminder_df['DATE'] < ptp_reminder_df['PTP_DATE_ORIG']]
            ptp_reminder = ptp_reminder_df['ACCOUNT NO.'].nunique()

            # PAYMENT REMINDER: Count reminders on or after PTP DATE for PTP accounts
            payment_reminder_df = reminder_df.merge(ptp_accounts, on='ACCOUNT NO.', how='inner')
            payment_reminder_df = payment_reminder_df[payment_reminder_df['DATE'] >= payment_reminder_df['PTP_DATE_ORIG']]
            payment_reminder = payment_reminder_df['ACCOUNT NO.'].nunique()

        confirmed_df = group[pd.to_numeric(group['CLAIM PAID AMOUNT'], errors='coerce') > 0]
        # Non-unique count for CONFIRMED
        confirmed = len(confirmed_df)
        confirmed_amount = pd.to_numeric(confirmed_df['CLAIM PAID AMOUNT'], errors='coerce').sum()

        # Adjust rates based on non-unique counts
        unique_worked_account = group[group['REMARK TYPE'] == 'Outgoing']['ACCOUNT NO.'].nunique()
        unique_rpc = rpc_df['ACCOUNT NO.'].nunique()
        unique_ptp = agent_ptp_df['ACCOUNT NO.'].nunique()
        unique_confirmed = confirmed_df['ACCOUNT NO.'].nunique()

        rpc_rate = round(unique_rpc / unique_worked_account * 100, 2) if unique_worked_account else 0
        ptp_rate = round(unique_ptp / unique_rpc * 100, 2) if unique_rpc else 0
        cvr = round(unique_confirmed / unique_ptp * 100, 2) if unique_ptp else 0

        summary_rows.append({
            'PERIOD': period,
            'AGENT': agent,
            'CAMPAIGN': campaign,
            'WORKED ACCOUNT': worked_account,
            'TOTAL CONNECTED': total_connected,
            'TOTAL TALK TIME': total_talk_time,
            'RPC': rpc,
            'RPC TALK TIME': rpc_talk_time,
            'PTP': ptp,
            'PTP TALK TIME': ptp_talk_time,
            'PTP REMINDER': ptp_reminder,
            'PAYMENT REMINDER': payment_reminder,
            'CONFIRMED': confirmed,
            'RPC RATE': f"{rpc_rate:.2f}%",
            'PTP RATE': f"{ptp_rate:.2f}%",
            'CVR': f"{cvr:.2f}%",
            'PTP AMOUNT': ptp_amount,
            'CONFIRMED AMOUNT': confirmed_amount
        })

    if summary_rows:
        summary_table = pd.DataFrame(summary_rows, columns=summary_columns)
    return summary_table.sort_values(by=['AGENT', 'CAMPAIGN'])

if uploaded_files := st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True):
    all_agent_productivity = []

    for idx, file in enumerate(uploaded_files, 1):
        df = load_data(file)
        # Debug: Print column names to verify
        st.write(f"Columns in file {idx}: {list(df.columns)}")
        df = process_file(df)
        df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
        df = df[df['DATE'].notna()]
        
        agent_productivity_summary = calculate_agent_productivity_summary(df)
        all_agent_productivity.append(agent_productivity_summary)

    agent_productivity_summary = pd.concat(all_agent_productivity, ignore_index=True).sort_values(by=['AGENT', 'CAMPAIGN'])

    summary_groups = {
        'Agent Productivity': {'Agent Productivity Summary': agent_productivity_summary}
    }

    st.write("## Agent Productivity Summary Table")
    st.write(agent_productivity_summary)

    st.download_button(
        label="Download Agent Productivity Summary as Excel",
        data=to_excel(summary_groups),
        file_name=f"Agent_Productivity_Summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
