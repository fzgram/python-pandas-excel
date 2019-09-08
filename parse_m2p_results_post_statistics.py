#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import os
import sys
import csv
import re
import matplotlib.pyplot as plt
import openpyxl


# In[2]:


def is_interactive():
    import __main__ as main
    return not hasattr(main, '__file__')


# ## Sheet 1 - all runs

# In[3]:


def add_all_runs_sheet(c_df, writer):
    c_df.to_excel(writer, sheet_name='All runs', index=False)    


# ## Sheet 2 - only non-outliers from all runs

# In[4]:


def add_runs_without_outliers(c_df, writer, outlier_threshold=120):
    no_outliers_df = c_df[c_df['AR_fps'] >= 120]
    no_outliers_df.to_excel(writer, sheet_name='Runs without outliers', index=False)
    return no_outliers_df


# ## Sheet 3 - only outliers from all runs

# In[5]:


def add_outlier_runs(c_df, writer, outlier_threshold=120):
    with_outliers_df = c_df[c_df['AR_fps'] < 120]
    with_outliers_df.to_excel(writer, sheet_name='Outlier runs', index=False)
    return with_outliers_df


# ## Sheet 4 - add statistics (min, max, stddev, etc.) for FPS (runs without outliers)

# In[6]:


def add_statistics(c_df, writer):
    sliced_df = c_df[['run', 'AR_fps']].copy()
    desc_df = sliced_df.describe()
    median_series = sliced_df.median()
    
    fps_stat_df = pd.DataFrame(columns=['Min', 'Max', 'Average', 'Median', 'Std Deviation'])
         
    fps_stat_df['Min'] = [desc_df["AR_fps"]["min"]]
    fps_stat_df['Max'] = [desc_df["AR_fps"]["max"]]
    fps_stat_df['Average'] = [desc_df["AR_fps"]["mean"]]
    fps_stat_df['Median'] = [median_series["AR_fps"]]
    fps_stat_df['Std Deviation'] = [desc_df["AR_fps"]["std"]]

    # Round to 4 decimals
    fps_stat_df = fps_stat_df.round(4)
    
    # Write the sliced_df to excel
    sliced_df.to_excel(writer, sheet_name='Statistics', index=False)
    
    #get a pointer to the same sheet to write other dfs and text to the same sheet
    curr_sheet = writer.sheets['Statistics']
    
    # Write text and fps_stat_df
    #curr_sheet.write(1, 4, "Statistics, # of Frames Delay")
    curr_sheet['E2'] = "Statistics of FPS values"
    fps_stat_df.to_excel(writer, startrow=2, startcol=4, sheet_name='Statistics', index=False)
    
    return fps_stat_df


# ## Sheet 5 - Analyze FPS values from all runs

# In[7]:


def fps_all_analysis(c_df, writer):
    curr_row = 0
    fps_col_series = c_df['AR_fps'].copy()
        
    # Convert the column to dataframe with unique values and their count
    fps_unique_count_df = fps_col_series.value_counts().sort_index().to_frame('Count of frame delay')
    fps_unique_count_df.rename_axis('Frame Delay', inplace=True)
    fps_unique_count_df.to_excel(writer, sheet_name='Frame Delay-Allruns', index=True)
    
    # Get current sheet pointer for future writing
    curr_sheet = writer.sheets['Frame Delay-Allruns']
    
    # Add grand total of runs
    curr_row = len(fps_unique_count_df) + 2 # update current row val. +2 because, 1 for column names of the dataframe and 1 for next empty row
    curr_sheet.cell(row=curr_row, column=1).value = 'Grand Total'
    curr_sheet.cell(row=curr_row, column=2).value = fps_unique_count_df.sum()[0]
    
    # Add outlier vs non-outlier stats
    curr_row = curr_row + 1 # update current row val
    
    outlier_stat_df = pd.DataFrame(columns=['Frames Delayed', 'Distribution of Frame Delay with all runs', 'in %'])
    total_non_outliers = len(fps_col_series[fps_col_series[0] <= 4])
    total_outliers = len(fps_col_series[fps_col_series[0] > 4])
    total_sum = fps_unique_count_df.sum()[0]

    outlier_stat_df.loc[0] = ['<=4', total_non_outliers, round((total_non_outliers*100)/total_sum, 2) ]
    outlier_stat_df.loc[1] = ['>4', total_outliers, round((total_outliers*100)/total_sum, 2) ]
    outlier_stat_df.loc[2] = ['total', total_sum, 100]
    outlier_stat_df.to_excel(writer, sheet_name='Frame Delay-Allruns', startrow=curr_row, index=False)
    
    # Add 3D pie chart image on the excel sheet
    data = outlier_stat_df['in %'].values.tolist()[:-1]
    labels = ['<=4', '>4']
    plt.title("Distribution of Frame Delay with all runs, in %'")
    plt.pie(data, labels=labels, autopct='%1.1f%%', startangle=120, shadow=True)
    piefile = f"{input_excel_name}_FrameDelayAnalysis.png"
    plt.savefig(piefile, dpi = 75)
    img = openpyxl.drawing.image.Image(piefile)
    img.anchor = 'G4'
    curr_sheet.add_image(img)
    
    plt.close('all')
    print(f"Saved pie chart: {piefile}")


# ## Sheet 6 - Analyze AR_fps column from all non-outlier runs

# In[8]:


def fps_no_outliers_analysis(no_outliers_df, writer):
    curr_row = 0
    fps_col_series = no_outliers_df['AR_fps'].copy()
    
    # Convert the column to dataframe with unique values and their count
    fps_unique_count_df = fps_col_series.value_counts().sort_index().to_frame()
    fps_unique_count_df.rename_axis('FPS unique values', inplace=True)
    fps_unique_count_df.rename(columns={'AR_fps':'count'})
    fps_unique_count_df['% of FPS value Distribution'] = round(fps_col_series.value_counts(normalize=True)*100, 2)
    fps_unique_count_df.sort_index()
    fps_unique_count_df.to_excel(writer, sheet_name='FPS-Runswithoutoutliers', index=True)
        
    # Get current sheet pointer for future writing
    curr_sheet = writer.sheets['FPS-Runswithoutoutliers']
    
    # Add grand total of runs
    curr_row = len(fps_unique_count_df) + 2 # update current row val
    curr_sheet.cell(row=curr_row, column=1).value = 'Grand Total'
    curr_sheet.cell(row=curr_row, column=2).value = fps_unique_count_df.sum()[0]
    curr_sheet.cell(row=curr_row, column=3).value = round(fps_unique_count_df['% of FPS value Distribution'].sum())
    
    
    # Add 3D pie chart image on the excel sheet
    data = fps_unique_count_df['% of FPS value Distribution'].values.tolist()
    labels = fps_unique_count_df.index.values.tolist()
    plt.title("Distribution of Frame Delay with non-outlier runs, in %'")
    plt.pie(data, labels=labels, autopct='%1.1f%%', startangle=120)
    piefile = f"{input_excel_name}_FrameDelayAnalysisNonOutliers.png"
    plt.savefig(piefile, dpi = 75)
    img = openpyxl.drawing.image.Image(piefile)
    img.anchor = 'G4'
    curr_sheet.add_image(img)
    
    plt.close('all')
    print(f"Saved pie chart: {piefile}")


# # MAIN

# In[9]:


#main
if is_interactive():
    input_excel = 'consolidation_result.xlsx'
else:
    input_excel = sys.argv[1]

# get the name of input excel file, discard the extension
input_excel_name, _ = os.path.splitext(input_excel)
excel_file = f'{input_excel_name}_post_analysis.xlsx'
writer = pd.ExcelWriter(excel_file, engine='openpyxl')

print(f"*** Working on folder: {input_excel} ***")

# Get the input excel sheet into a dataframe
c_df = pd.read_excel(input_excel, 0, index_col=None)


# In[10]:


##### Add required sheets #######

# Sheet 1 - all runs
print("Working on Sheet 1 - All runs")
add_all_runs_sheet(c_df, writer)
print(f"Total runs: {len(c_df)} ")
print("DONE!\n")

# Sheet 2 - only non-outliers from all runs
print("Working on Sheet 2 - only non-outliers from all runs")
no_outliers_df = add_runs_without_outliers(c_df, writer)
print(f"Total non-outlier runs: {len(no_outliers_df)} ")
print("DONE!\n")

# Sheet 3 - only outliers from all runs
print("Working on Sheet 3 - only ourtliers from all runs")
with_outliers_df = add_outlier_runs(c_df, writer)
print(f"Total outlier runs: {len(with_outliers_df)} ")
print("DONE!\n")

# Sheet 4 - add statistics (min, max, stddev, etc.) for frame delay (runs without outliers)
print("Working on Sheet 4 - add statistics (min, max, stddev, etc.) for 'AR_fps' values for non-outlier runs")
fps_stat_df = add_statistics(no_outliers_df, writer)
print("DONE!\n")

# Sheet 5 - Analyze frame delay column from all runs
print("Working on Sheet 5 - Analyze frame delay column from all runs")
#fps_all_analysis(c_df, writer)
print("DONE!\n")

# Sheet 6 - Analyze frame delay distribution of non-outlier runs
print("Working on Sheet 6 - Analyze frame delay distribution of non-outlier runs")
fps_no_outliers_analysis(no_outliers_df, writer)
print("DONE!\n")

# Final step. Save the Excel writer object and close it
print(f"Consolidating all sheets in final Excel: {excel_file}")
writer.save()
writer.close()
print("DONE!")


# In[ ]:




