import pandas as pd

def df_to_list(df, row=None):
	if row is None:
		return map(list, df.values)
	else:
		return map(list, df.values)[row]

def create_sheetname(df, cols):
	return " ".join(df_to_list(df[cols], row=0))

xlfile = "DS O Block Rev C.xlsx"
master_id = 12

master_info_cols = ['ControllerName', 'MasterType']
slave_info_cols = ['AssocMaster', 'SlaveType']
commission_cols = ['IOConnection', 'PointType', 'System', 'Description', 'WireNumber']
commission_rename_cols = ['IO', 'Type', 'System', 'Description', 'Wire#']
extra_commission_cols = ["Pass/Fail", "Signed", "Date"]

df = pd.read_excel(xlfile, sheetname='Sheet1')
xenta_df = df[df['MasterID29'] == master_id].copy()

workbookname = create_sheetname(xenta_df, master_info_cols) + '.xlsx'
writer = pd.ExcelWriter(workbookname, engine='xlsxwriter')

slave_ids = xenta_df['SlaveID30'].unique()

slave_dfs = []

for slave_id in slave_ids:
	# create unique df for each slave module df
	# append slave module df to list
	slave_dfs.append(xenta_df[xenta_df['SlaveID30'] == slave_id])


print(df[commission_cols].head())

print(xenta_df[master_info_cols].head())

print(workbookname)

for slave_df in slave_dfs:
	sheetname = create_sheetname(slave_df, slave_info_cols)
	commission_sheet = slave_df[commission_cols]
	# rename columns for Excel spreadsheet
	commission_sheet.columns = commission_rename_cols
	# added xtra sign off columns for spreadsheet
	commission_sheet = pd.concat([commission_sheet, pd.DataFrame(columns=extra_commission_cols)])
	# reorder columns, the above concat orders cols alphabetically
	commission_sheet = commission_sheet[commission_rename_cols + extra_commission_cols]
	# add df as new sheet in Excel workbook	
	commission_sheet.to_excel(writer, sheet_name=sheetname, index=False)

writer.save()

# MasterID29
# SlaveID30
# InUse
# IsPreDefined
# IOConnection
# Description
# PointType
# WireNumber
# OrderByPoint
# CableName
# DeviceName
# ActualDeviceName
# DeviceManufacturer
# DeviceCost
# System
# AssocMaster
# MasterType
# SlaveType
# ControllerName
# Comments31
