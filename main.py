import os.path

import pandas as pd
from openpyxl import load_workbook

TICKETS_RAW_DATA = 'test_files/Ticket Search Results (2).csv'
DESTINATION_SHEET = 'test_files/'


def main():
    """Main code block."""
    ticket_df = pd.read_csv(TICKETS_RAW_DATA)
    file_name = 'test_01.xlsx'
    issue_summary = get_count(ticket_df, 'Issue Type')
    resource_summary = get_count(ticket_df, 'Resources')
    total_time_worked = ticket_df['Total Hours Worked'].sum()
    closed_overdue_tickets = get_closed_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    open_overdue_tickets = get_open_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    all_overdue_tickets = get_all_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    #   TURN LIST INTO DATAFRAME
    overdue_df = get_overdue_summary(ticket_df, all_overdue_tickets)
#   Write to EXCEL
    with pd.ExcelWriter(file_name) as writer:
        ticket_df.to_excel(writer, sheet_name='raw_data')
        issue_summary.to_excel(writer, sheet_name='summary', startcol=0)
        resource_summary.to_excel(writer, sheet_name='summary', startcol=4)
        overdue_df.to_excel(writer, sheet_name='summary', startcol=8)


def get_overdue_summary(dataframe, overdue_tickets):
    """Display all overdue tickets' details."""
    # print(overdue_tickets)
    overdue_detail = dataframe[dataframe['Ticket Number'].isin(overdue_tickets)]
    # pd.Dataframe(overdue_detail[['Title', 'Due', 'Resolved Time', 'Status']])
    return pd.DataFrame(overdue_detail[['Title', 'Due', 'Resolved Time', 'Status']])


def get_count(dataframe, column):
    """Get total count of entries in column."""
    return dataframe[column].value_counts(dropna=False)


def get_sum(dataframe, column):
    """Get total sum of entries in column."""
    return dataframe[column].sum


def get_all_overdue_tickets(dataframe, due_date, resolved):
    """Get a list of overdue tickets."""
    resolved_column = pd.to_datetime(dataframe[resolved], dayfirst=True).dt.day
    # resolved_date = resolved_column.dt.day
    due_column = pd.to_datetime(dataframe[due_date], dayfirst=True).dt.day
    # due_date_dates = due_column.dt.day
    # difference = due_date_dates - resolved_date
    difference = due_column - resolved_column
    # print(difference)
    overdue_tickets = dataframe.where((due_column - resolved_column < 0)).dropna()['Ticket Number']
    overdue_ticket_numbers = sorted(i for i in overdue_tickets)
    # return overdue_ticket_numbers
    return overdue_tickets


def get_open_overdue_tickets(dataframe, due_date, resolved):
    """Get a list of overdue tickets."""
    resolved_column = pd.to_datetime(dataframe[resolved], dayfirst=True).dt.day
    # resolved_date = resolved_column.dt.day
    due_column = pd.to_datetime(dataframe[due_date], dayfirst=True).dt.day
    # due_date_dates = due_column.dt.day
    # difference = due_date_dates - resolved_date
    difference = due_column - resolved_column
    # print(difference)
    overdue_tickets = dataframe.where((due_column - resolved_column < 0) & (dataframe['Status'] != 'Complete')).dropna()['Ticket Number']
    overdue_ticket_numbers = sorted(i for i in overdue_tickets)
    return overdue_ticket_numbers
    # return overdue_tickets


def get_closed_overdue_tickets(dataframe, due_date, resolved):
    """Get a list of overdue tickets."""
    resolved_column = pd.to_datetime(dataframe[resolved], dayfirst=True).dt.day
    # resolved_date = resolved_column.dt.day
    due_column = pd.to_datetime(dataframe[due_date], dayfirst=True).dt.day
    # due_date_dates = due_column.dt.day
    # difference = due_date_dates - resolved_date
    difference = due_column - resolved_column
    # print(difference)
    overdue_tickets = dataframe.where((due_column - resolved_column < 0) & (dataframe['Status'] == 'Complete')).dropna()['Ticket Number']
    overdue_ticket_numbers = sorted(i for i in overdue_tickets)
    return overdue_ticket_numbers
    # return overdue_tickets


if __name__ == '__main__':
    main()
