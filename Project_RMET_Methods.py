import xlrd
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import random

def select_random_sheets(num_sheets, total_sheets, with_replacement):
    sequence = range(1,total_sheets+1)
    if with_replacement == True:
        sample = np.random.choice(sequence, num_sheets)
    else:
        sample = random.sample(sequence, num_sheets)
    return sample

def consolidate_sample_sheets(filename, sample):
    sheetname = "Student " + str(sample[0])

    df = pd.read_excel(filename, sheet_name=sheetname, usecols='M:Q')
    for index in range(1, len(sample)):
        sheetname = "Student " + str(sample[index])
        df_temp = pd.read_excel(filename, sheet_name=sheetname, usecols='M:Q')
        df = df + df_temp
    df.columns = ['Timestamp', 'Difficult', 'Easy', 'Boring', 'Engaging']
    df['Timestamp'] = df['Timestamp']//len(sample)
    df = df.dropna()
    return df

def running_average_two_rows(data_frame):
    number_of_rows = len(data_frame.index)
    lastdict = {'Timestamp':[number_of_rows], 'Difficult':[0], 'Easy':[0], 'Boring':[0], 'Engaging':[0]   }
    last_frame = pd.DataFrame(lastdict)
    data_frame = pd.concat([data_frame, last_frame], ignore_index = True)
    for index in range(len(data_frame)-1):
        data_frame.iloc[index] = ((data_frame.iloc[index] + data_frame.iloc[index+1]))/2
        data_frame.iloc[index] = (data_frame.iloc[index]).apply(np.ceil)
    data_frame['Timestamp'] = data_frame['Timestamp'] -1
    data_frame = data_frame.drop(len(data_frame)-1)
    data_frame['Timestamp'][number_of_rows-1] = data_frame['Timestamp'][number_of_rows-1] + 1
    return data_frame

def calculate_net_percentages(sample_dataframe, number_of_sheets):
    number_of_rows = len(sample_dataframe.index)
    dict_temp = {'Timestamp': list(range(0, number_of_rows)), 'Net Engagement':[0]*(number_of_rows), 'Net Difficulty':[0]*(number_of_rows)}
    net_value_data_frame = pd.DataFrame(dict_temp)
    net_value_data_frame['Net Engagement'] = sample_dataframe['Engaging'] - sample_dataframe['Boring']
    net_value_data_frame['Net Difficulty'] = sample_dataframe['Difficult'] - sample_dataframe['Easy']
    net_value_data_frame['Net Engagement'] = round((net_value_data_frame['Net Engagement']/number_of_sheets)*100,2)
    net_value_data_frame['Net Difficulty'] = round((net_value_data_frame['Net Difficulty']/number_of_sheets)*100,2)
    # net_value_data_frame['Net Engagement'] = net_value_data_frame['Net Engagement'].astype('int')
    # net_value_data_frame['Net Difficulty'] = net_value_data_frame['Net Difficulty'].astype('int')
    return net_value_data_frame

def calculateSSE(net_value_percent_population, net_value_percent_sample):
    number_of_rows = len(net_value_percent_sample.index)
    dict_temp = {'Timestamp': list(range(0, number_of_rows)), 'Net Engagement':[0]*(number_of_rows), 'Net Difficulty':[0]*(number_of_rows)}
    error_data_frame = pd.DataFrame(dict_temp)
    error_data_frame['Net Engagement'] = net_value_percent_population['Net Engagement'] - net_value_percent_sample['Net Engagement']
    error_data_frame['Net Difficulty'] = net_value_percent_population['Net Difficulty'] - net_value_percent_sample['Net Difficulty']
    error_data_frame['Net Engagement'] = error_data_frame['Net Engagement']**2
    error_data_frame['Net Difficulty'] = error_data_frame['Net Difficulty']**2
    net_engagement_sse = error_data_frame['Net Engagement'].sum()
    net_difficulty_sse = error_data_frame['Net Difficulty'].sum()
    return net_engagement_sse, net_difficulty_sse