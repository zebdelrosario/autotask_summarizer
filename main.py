import os.path
import numpy as np
import pandas as pd
from openpyxl import load_workbook

RAW_DATA = input('Input file name:\n>')
DESTINATION_SHEET = 'test_files/'


def main():
    """Main code block."""
    tickets_raw_data = f'test_files/{RAW_DATA}.csv'
    ticket_df = pd.read_csv(tickets_raw_data)
    ticket_resources = get_unique_entries(ticket_df, 'Resources')
    issue_summary = get_count(ticket_df, 'Issue Type')
    resource_summary = get_count(ticket_df, 'Resources')
    total_time_worked = ticket_df['Total Hours Worked'].sum()
    closed_overdue_tickets = get_closed_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    open_overdue_tickets = get_open_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    all_overdue_tickets = get_all_overdue_tickets(ticket_df, 'Due', 'Resolved Time')
    overdue_ticket_numbers = get_all_overdue_ticket_numbers(ticket_df, 'Due', 'Resolved Time')
    #   TURN LIST INTO DATAFRAME
    overdue_df = get_overdue_summary(ticket_df, all_overdue_tickets)
#     Get resource time summary
    resource_time_summary = get_unique_summary(ticket_df, 'Resources', 'Total Hours Worked')
    # print(resource_summary)
    complete_resource_summary = pd.concat([resource_time_summary, resource_summary.rename('Tickets Assigned')], axis=1)
    print('Preparing Data...')
#   Sheet Styling
    highlighted_rows = ticket_df['Ticket Number'].isin(overdue_ticket_numbers).map({
        True: 'background-color: yellow',
        False: ''
    })
    styler = ticket_df.style.apply(lambda _: highlighted_rows)
#   Write to EXCEL
    out_file_name = input('Output file name:\n>') + '.xlsx'
    with pd.ExcelWriter(out_file_name) as writer:
        ticket_df.to_excel(writer, sheet_name='raw_data')
        print('Preparing File')
        styler.to_excel(writer, sheet_name='raw_data')
        print('Highlighting cells...')
        issue_summary.to_excel(writer, sheet_name='summary', startcol=0)
        print('Writing Issue Types...')
        complete_resource_summary.to_excel(writer, sheet_name='summary', startcol=4)
        print('Writing Resource summary...')
        overdue_df.to_excel(writer, sheet_name='summary', startcol=8)
        print('Writing Overdue tickets...')


def get_unique_summary(dataframe, column_one, column_two):
    """Summarize column based on unique entries."""
    summary = dataframe.groupby(column_one)[column_two].sum()
    return summary


def get_unique_entries(dataframe, column):
    """Get a list of resources."""
    column_to_dissect = dataframe[column]
    unique_entries = []
    for i in column_to_dissect:
        if i not in unique_entries:
            unique_entries.append(i)
    return unique_entries


def get_overdue_summary(dataframe, overdue_tickets):
    """Display all overdue tickets' details."""
    # print(overdue_tickets)
    overdue_detail = dataframe[dataframe['Ticket Number'].isin(overdue_tickets)]
    # pd.Dataframe(overdue_detail[['Title', 'Due', 'Resolved Time', 'Status']])
    return pd.DataFrame(overdue_detail[['Title', 'Due', 'Resolved Time', 'Status']])


def get_count(dataframe, column):
    """Get total count of entries in column."""
    return dataframe[column].value_counts(dropna=False).sort_index()


def get_sum(dataframe, column):
    """Get total sum of entries in column."""
    return dataframe[column].sum


def get_all_overdue_ticket_numbers(dataframe, due_date, resolved):
    """Get a list of overdue ticket numbers."""
    resolved_column = pd.to_datetime(dataframe[resolved], dayfirst=True).dt.day
    due_column = pd.to_datetime(dataframe[due_date], dayfirst=True).dt.day
    difference = due_column - resolved_column
    overdue_tickets = dataframe.where((due_column - resolved_column < 0)).dropna()['Ticket Number']
    overdue_ticket_numbers = sorted(i for i in overdue_tickets)
    return overdue_ticket_numbers


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
