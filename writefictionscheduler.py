#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 15 06:23:45 2021

@author: dfox
"""

import pandas as pd
import datetime
import os

currentdate = datetime.date.today()
currenttime = datetime.datetime.now()

out_location = "."
in_location = "."
outputfile = os.path.join(out_location,
    "WFSOutput_{:4d}{:02d}{:02d}{:02d}{:02d}{:02d}.xlsx" \
        .format(currenttime.year, currenttime.month, currenttime.day,
                currenttime.hour, currenttime.minute, currenttime.second))
inputfile = os.path.join(in_location, "WFSInputExample.xlsx")

non_work_cols = ['WORK', 'PRIORITY', 'TYPE', 'SIZE', 'BETA START',
                 'EDITOR START', 'PROOF START', 'REVIEW START',
                 'RELEASE DATE']
non_period_cols = ['WORK', 'TYPE', 'SIZE', 'SUMELSE']
non_period_rows = ['MKTG', 'ADMIN', 'SUMELSE']
periods_cols = ['DATE', 'WORKING', 'FD', 'MKTG', 'ADMIN', 'REST', '@OTHERS',
                'NOTES', 'COMPLETED', 'WORKING - ACTUAL', 'FD - ACTUAL', 'MKTG - ACTUAL',
                'ADMIN - ACTUAL', 'REST - ACTUAL']

def increment_cell(df, index, col, value):
    if col not in df:
        df[col] = 0
    elif type(df.at[index, col]) != int:
        df[col].astype(int)
    df.at[index, col] = df.at[index, col] + value

def set_cell(df, index, col, value):
    if col not in df:
        df[col] = 0
    df.at[index, col] = value

# set initial state

def main_pass(wadf, padf, index, fd_to_allocate, rest_to_allocate, 
              others_to_allocate):
    # populate plan needed for first draft first
    for windex, wrow in wadf.iterrows():
        for col in wadf:
            if col in non_work_cols:
                continue
            if wrow[col] <= 0:
                continue
            if col == 'PLAN':
                if rest_to_allocate > 0:
                    plan_needed = wadf['PLAN'][windex]
                    if plan_needed > rest_to_allocate:
                        # plan takes up rest plus eats into fd
                        col_name = wrow['WORK'] + ' PLAN'
                        increment_cell(padf, index, col_name, rest_to_allocate)
                        plan_needed = plan_needed - rest_to_allocate
                        set_cell(wadf, windex, 'PLAN', plan_needed)
                        rest_to_allocate = 0
                    elif plan_needed > 0:
                        # plan takes up some of rest
                        col_name = wrow['WORK'] + ' PLAN'
                        increment_cell(padf, index, col_name, plan_needed)
                        rest_to_allocate = rest_to_allocate - plan_needed
                        plan_needed = 0
                        set_cell(wadf, windex, 'PLAN', plan_needed)
            elif col == 'FD':
                if fd_to_allocate > 0:
                    fd_needed = wadf['FD'][windex]
                    if fd_needed > fd_to_allocate:
                        col_name = wrow['WORK'] + ' FD'
                        increment_cell(padf, index, col_name, fd_to_allocate)
                        increment_cell(wadf, windex, 'FD', -fd_to_allocate)
                        fd_to_allocate = 0
                    else:
                        col_name = wrow['WORK'] + ' FD'
                        increment_cell(padf, index, col_name, fd_needed)
                        set_cell(wadf, windex, 'FD', 0)
                        fd_to_allocate = fd_to_allocate - fd_needed
            elif col.startswith("@"):
                if others_to_allocate > 0:
                    col_name = wrow['WORK'] + " " + col
                    others_needed = wrow[col]
                    if others_needed > others_to_allocate:
                        increment_cell(padf, index, col_name, 
                                       others_to_allocate)
                        increment_cell(wadf, windex, col, -others_to_allocate)
                    else:
                        increment_cell(padf, index, col_name, others_needed)
                        set_cell(wadf, windex, col, 0)
            elif rest_to_allocate > 0:
                col_name = wrow['WORK'] + " " + col
                task_needed = wrow[col]
                if task_needed > rest_to_allocate:
                    increment_cell(padf, index, col_name, rest_to_allocate)
                    increment_cell(wadf, windex, col, -rest_to_allocate)
                    rest_to_allocate = 0
                else:
                    increment_cell(padf, index, col_name, task_needed)
                    set_cell(wadf, windex, col, 0)
                    rest_to_allocate = rest_to_allocate - task_needed
            break
    return fd_to_allocate

def plan_needed_pass(wadf, padf, index, fd_to_allocate):
    # populate plan needed for first draft first
    for windex, wrow in wadf.iterrows():
        if fd_to_allocate > 0:
            for col in wadf:
                if col == 'WORK' or col == 'TYPE' or col == 'SIZE':
                    continue
                if wrow[col] <= 0:
                    continue
                if col == 'PLAN' and fd_to_allocate > 0:
                    plan_needed = wadf['PLAN'][windex]
                    if plan_needed > fd_to_allocate:
                        # plan takes up rest plus eats into fd
                        col_name = wrow['WORK'] + ' PLAN'
                        increment_cell(padf, index, col_name, fd_to_allocate)
                        plan_needed = plan_needed - fd_to_allocate
                        set_cell(wadf, windex, 'PLAN', plan_needed)
                        fd_to_allocate = 0
                    elif plan_needed > 0:
                        # plan takes up some of rest
                        col_name = wrow['WORK'] + ' PLAN'
                        increment_cell(padf, index, col_name, plan_needed)
                        fd_to_allocate = fd_to_allocate - plan_needed
                        plan_needed = 0
                        set_cell(wadf, windex, 'PLAN', plan_needed)
                break
        if fd_to_allocate <= 0:
            break

def gather_stats(wadf, sdf, padf, paindex):
    for index, row in sdf.iterrows():
        if row['WORK'] in non_period_rows:
            continue
        for col in sdf:
            if col not in non_period_cols:
                if row[col] > 0:
                    col_name = row['WORK'] + " " + col
                    wb = row[col]
                    set_cell(padf, paindex, col_name, wb)
                    wawdf = wadf[wadf['WORK'] == row['WORK']]
                    windex = wawdf.index[0]
                    increment_cell(wadf, windex, col, -wb)
                    
#    spbdf = sdf[(sdf['WORK'] not in non_period_rows)]
    spbdf = sdf[sdf['WORK'] != "SUMELSE"]
    fdwb = spbdf['FD'].sum()
    wwb = 0
    for col in spbdf:
#        if col not in non_period_cols and col != "FD" and not col.startswith("@"):
        if col not in non_period_cols and not col.startswith("@"):
           wwb = wwb + spbdf[col].sum()
    spmdf = sdf[sdf['WORK'] == 'MKTG']
    mwb = spmdf.iloc[0]['PLAN']
    samdf = sdf[sdf['WORK'] == 'ADMIN']
    awb = samdf.iloc[0]['PLAN']
    rwb = wwb - fdwb - mwb - awb
    set_cell(padf, paindex, 'MKTG - ACTUAL', mwb)
    set_cell(padf, paindex, 'ADMIN - ACTUAL', awb)
    set_cell(padf, paindex, 'FD - ACTUAL', fdwb)
    set_cell(padf, paindex, 'REST - ACTUAL', rwb)
    set_cell(padf, paindex, 'WORKING - ACTUAL', wwb)

xl = pd.ExcelFile(inputfile)
sbdict = xl.parse(sheet_name=None)

wdf = sbdict['WORKS']
wadf = wdf.copy()
wadf.sort_values('PRIORITY', inplace=True)
pdf = sbdict['PERIODS']
padf = pdf.copy()

padf.sort_values('DATE')
padf.reset_index(inplace=True, drop=True)
"""
pandf = padf[padf['DATE'] > currenttime]
panindex = pandf.index[0]
if panindex != 0:
    pafindex = panindex - 1
else:
    pafindex = 0
"""
pandf = padf[padf['COMPLETED'] == False]
pafindex = pandf.index[0]
    
for index, row in padf.iterrows():
    if index < pafindex:
        sdate = row['DATE']
        sdf = sbdict['PERIOD {y:4d}.{m:02d}.{d:02d}'.format(y=sdate.year, 
                                                            m=sdate.month, 
                                                            d=sdate.day)]
        gather_stats(wadf, sdf, padf, index)
    else:
        fd_to_allocate = main_pass(wadf, padf, index, row['FD'], row['REST'],
                                   row['@OTHERS'])
        if fd_to_allocate > 0:
            plan_needed_pass(wadf, padf, index, fd_to_allocate)

for index, row in wadf.iterrows():
    for col in wadf:
        if col.startswith("@"):
            col_name = " ".join([row['WORK'], col])
            if col_name in padf:
                pardf = padf[padf[col_name] > 0]
                rows, cols = pardf.shape
                if rows > 0:
                    sdate = pardf.iloc[0]['DATE']
                    warddf = wadf[wadf['WORK'] == row['WORK']]
                    rows, cols = warddf.shape
                    if rows > 0:
                        wcol_name = " ".join([col[1:], "DATE"])
                        wadf.at[warddf.index[0], wcol_name] = sdate

wpdf = wdf.sort_values('PRIORITY')
for windex, wrow in wpdf.iterrows():
    for col in wpdf:
        if col in non_work_cols:
            continue
        if wrow[col] == 0:
            continue
        col_name = " ".join([wrow['WORK'], col])
        if col_name in padf:
            periods_cols.append(col_name)
            
padf = padf[periods_cols]

works_cols = []
for col in wdf:
    works_cols.append(col)
for col in wdf:
    if col.startswith("@"):
        works_cols.append(" ".join([col[1:], "DATE"]))
wadf = wadf[works_cols]
           
writer = pd.ExcelWriter(outputfile, 
                        date_format='M/D/YYY', datetime_format='M/D/YYYY',
                        engine='xlsxwriter')

padf.to_excel(writer, sheet_name="PERIODS", index=False)
wadf.to_excel(writer, sheet_name="WORKS", index=False)

# Next enhancement
#   Create PERIOD worksheets for upcoming periods
#   Rows work items
#   Columns are work items and number of blocks
max_period_sheets = 10
num_period_sheets = 0
periods_ignore_cols = ['DATE', 'WORKING', 'FD', 'REST', '@OTHERS', 'NOTES', 
                       'COMPLETED']
period_cols = ['Task', 'Blocks']
for index, row in pandf.iterrows():
    num_period_sheets += 1
    if num_period_sheets <= max_period_sheets:
        pdate = row['DATE']
        paddf = pd.DataFrame()
        for col in padf:
            if col not in periods_ignore_cols:
                workblocks = padf.at[index, col]
                if workblocks > 0:
                    periodrec = [col, workblocks]
                    periodrecseries = pd.Series(periodrec, period_cols)
                    paddf = paddf.append(periodrecseries, ignore_index=True)
        paddf = paddf[period_cols]
        sheet = 'PERIOD {y:4d}.{m:02d}.{d:02d}'.format(y=pdate.year, 
                                                       m=pdate.month, 
                                                       d=pdate.day)
        paddf.to_excel(writer, sheet_name=sheet, index=False)
        
# Next enhancement
#   Create a RELEASES worksheet with the dates of all releases

# Need to parse padf dataframe and filter out columns that have RELEASE in them
# but not @

milestones_dict = {}
release_columns = [col for col in padf if ' RELEASE' in col]
fd_columns = [col for col in padf if ' FD' in col]

# print(release_columns)
# print(fd_columns)

xpadf = padf[['DATE'] + release_columns]
# print(xpadf)
for col in release_columns:
    work = col[:-8]
    # print(work)
    # print(xpadf[col])
    xpadf = padf[padf[col] > 0]
    # print(xpadf)
    rows, cols = xpadf.shape
    if rows > 0:
        if work not in milestones_dict:
            milestones_dict[work] = {'FD START' : "",
                                     'FD COMPLETE' : "",
                                     'RELEASE' : ""}
        milestones_dict[work]['RELEASE'] = (xpadf.iloc[-1]['DATE']
                                            + datetime.timedelta(days=15))
xpadf = padf[['DATE'] + fd_columns]
for col in fd_columns:
    work = col[:-3]
    # print(work)
    xpadf = padf[padf[col] > 0]
    # print(xpadf)
    rows, cols = xpadf.shape
    if rows > 0:
        if work not in milestones_dict:
            milestones_dict[work] = {'FD START' : "",
                                     'FD COMPLETE' : "",
                                     'RELEASE' : ""}
        milestones_dict[work]['FD START'] = xpadf.iloc[0]['DATE']
        milestones_dict[work]['FD COMPLETE'] = (xpadf.iloc[-1]['DATE']
                                                + datetime.timedelta(days=13))
mdf = pd.DataFrame()
for work, workdict in milestones_dict.items():
    # print(workdict)
    dwrec = [work, workdict['FD START'], workdict['FD COMPLETE'], 
             workdict['RELEASE']]
    dwrecseries = pd.Series(dwrec, ['WORK', 'FD START', 'FD COMPLETE',
                                    'RELEASE'])
    mdf = mdf.append(dwrecseries, ignore_index=True)
mdf = mdf[['WORK', 'FD START', 'FD COMPLETE', 'RELEASE']]
mdf.to_excel(writer, sheet_name='MILESTONES', index=False)
           
writer.save()
print("wrote output file " + outputfile)