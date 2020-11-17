'''
Script I made to filter database things
'''
 
from typing import Dict, List
import argparse
import pandas as pd
import datetime as dt
from openpyxl import Workbook


MGR_LVL_EMPLOYEES = Dict[str, pd.DataFrame]
 
parser = argparse.ArgumentParser(
        description='Stack Ranking MyHome Sales Spreadsheets')
 
parser.add_argument('--input_file', '-i', type=str,
                    help='The path to the input xls.')
parser.add_argument('--sheet_name', '-s', type=str, default='Sheet1',
                    help='The name of the sheet with the records.')

# a list that will be used to create a dictionary of GMs for a region
# you only need to edit this if GM's change
# order isn't too relevant, but make sure that North, Central, and South are
# sequential so that you can easily screenshot them together
GM_LIST = ['NotReal Person','AnotherNotReal Person']

# exception dict, 'Employee Name": "Manager Name" for use when employees are
# filed under incorrect GM
MGR_REMAP = {
        'Joe Smith': 'Bob Fisher',
}


def load_data(filename: str, sheetname: str) -> pd.DataFrame:
    """Load a preprocessed input data frame.
 
   Replaces "No Activity" offer values with 0s.
 
   Args:
       filename: The relative of absolute filepath.
       sheetname: The name of the worksheet with the data.
   Returns:
       The loaded dataframe with columns 'Mgr', 'Name', 'Offers' and 'Accepts'
   """
    data = pd.read_excel(filename, sheet_name=sheetname, skiprows=5)
    data = data[['Mgr', 'Name', 'MTD\n Offers', 'MTD Accept']]
    data.columns = ['Mgr', 'Name', 'Offers', 'Accepts']
    data['Offers'][data['Offers'] == 'No Activity'] = 0
    return data
 
 
def replace_managers(
        data: pd.DataFrame, new_managers: Dict[str, str]) -> pd.DataFrame:
    """Replaces misattributed managers.
 
   Args:
       data: The dataframe with an employee 'Name' column  and 'Mgr' Column.
       new_managers: A dictionary mapping employees to actual managers.
   Returns:
       A Dataframe with a new "Mgr_actual" column.
   """
    remap = pd.DataFrame([
        {"Name": emp, "Mgr": mgr}
        for emp, mgr
        in new_managers.items()
    ])
    remapped = data.set_index('Name').join(
            remap.set_index('Name'), how="left", rsuffix="_actual")
    remapped.Mgr_actual.fillna(remapped.Mgr, inplace=True)
    return remapped.reset_index()
 
 
def split_by_manager(
        data: pd.DataFrame, managers:
        List[str] = None) -> MGR_LVL_EMPLOYEES:
    """Splits the offers/accepts data by manager into a Dataframe dictionary.
 
   Args:
       data: The data with columns Mgr_actual
       managers: A list with names of managers we want to track.
 
   Returns:
       A dictionary with managers as keys and dataframes with employees
       reporting to them.
   """
    managers = managers or GM_LIST
    managers = pd.DataFrame({'Mgr_actual': managers})
    subset = data.set_index('Mgr_actual').join(
            managers.set_index('Mgr_actual'), how='inner')
    subset.reset_index(inplace=True)
    by_manager = {
            mgr: subset[subset['Mgr_actual'] == mgr]
            for mgr in managers['Mgr_actual']
    }
    return by_manager
 

def prepare_for_excel(data: pd.DataFrame, mgr: str) -> pd.DataFrame:
    """Prepares a manager level data frame to be written to the trackin sheet.
 
   Args:
       data: The dataframe of employees reporting to a manager.
       mgr: The manage the employees report to. Will become the header of Name.
   Returns:
       A dataframe that can be written to the spreadsheet.
   """
    data = data[['Name', 'Offers', 'Accepts']]
    data.columns = [mgr, 'Offers', 'Accepts']
    row1 = pd.DataFrame([{
        mgr: '',
        'Offers': data['Offers'].sum(),
        'Accepts': data['Accepts'].sum()
    }])
    row1.iloc[0][0].set
    
   
    return pd.concat([row1, data])

def write_summary_to_spreadsheet(data: MGR_LVL_EMPLOYEES) -> None:
    """Write the manager-employees dictionary to a spreadsheet
 
   Args:
       data: A dictionary mapping managers to employee level data.
   """
    book = Workbook()
    filename = 'stack_ranker {}.xlsx'.format(dt.date.today().strftime('%Y %m %d'))
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    writer.book = book
    dfs = [prepare_for_excel(df, mgr) for mgr, df in data.items()]
    for i, df in enumerate(dfs):
        df.to_excel(writer, startcol = i * (df.shape[1] + 1), index=False)
    writer.save()

 
def main(args):
    data = load_data(args.input_file, args.sheet_name)
    data = replace_managers(data, MGR_REMAP)
    data = split_by_manager(data)
    write_summary_to_spreadsheet(data)
 
 
if __name__ == "__main__":
    main(parser.parse_args())