import xlrd
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import random
import Project_RMET_Methods as methods
import time

if __name__=="__main__":

    filename = "Segmentation_Paging.xlsx"
    totalnum = 69
    
    total_dataframe = methods.consolidate_sample_sheets(filename, range(1, totalnum+1))  
    total_dataframe = methods.running_average_two_rows(total_dataframe)
    net_dataframe_percent = methods.calculate_net_percentages(total_dataframe, totalnum)
    print(net_dataframe_percent)

    fig, axs = plt.subplots(figsize=(6.0, 4.0))
    net_dataframe_percent.plot('Timestamp', ['Net Engagement', 'Net Difficulty'], ax=axs)
    # plt.show()              
    fig.savefig("plot69.png")

    
    
    # combined_sample_data_frame = pd.DataFrame(columns=['Timestamp', 'Net Engagement', 'Net Difficulty'])
    # engagement_data_frame_compare = pd.DataFrame({'Timestamp': list(range(1, 49))})
    # difficulty_data_frame_compare = pd.DataFrame({'Timestamp': list(range(1, 49))})
    SSE_dataframe = pd.DataFrame(columns=('number_of_samples', 'net_engagement_sse', 'net_difficulty_sse'))
    number_of_iterations = 100

    time_initial = time.time()

    for samplenum in range(10,11, 5):
        engagement_data_frame_compare = pd.DataFrame({'Timestamp': list(range(1, 9))})
        difficulty_data_frame_compare = pd.DataFrame({'Timestamp': list(range(1, 9))})
        average_sse_engagement, average_sse_difficulty = 0, 0
        
        for i in range(number_of_iterations):
            t0 = time.time()
            sample = methods.select_random_sheets(samplenum, totalnum, False)
            sample_dataframe = methods.consolidate_sample_sheets(filename, sample)
            sample_dataframe = methods.running_average_two_rows(sample_dataframe)
            net_sample_dataframe_percent = methods.calculate_net_percentages(sample_dataframe, samplenum)

            engagement_data_frame_compare['Engagement' + str(i)] = net_sample_dataframe_percent['Net Engagement'].copy()
            difficulty_data_frame_compare['Difficulty' + str(i)] = net_sample_dataframe_percent['Net Difficulty'].copy()
            sse_engagement, sse_difficulty = methods.calculateSSE(net_dataframe_percent, net_sample_dataframe_percent)
            average_sse_engagement = average_sse_engagement + sse_engagement
            average_sse_difficulty = average_sse_difficulty + sse_difficulty
            SSE_dataframe.loc[len(SSE_dataframe.index)] = [samplenum, sse_engagement, sse_difficulty]


            print("Num_samples = " + str(samplenum) + "; Num_iteration = " + str(number_of_iterations) + "; Time = " + str(time.time() - t0))
        average_sse_difficulty = int(average_sse_difficulty/number_of_iterations)
        average_sse_engagement = int(average_sse_engagement/number_of_iterations)
        
        fig, (axs1, axs2) = plt.subplots(1, 2, figsize=(12.0, 4.0))
        cols1 = list(engagement_data_frame_compare.columns)
        cols2 = list(difficulty_data_frame_compare.columns)
        cols1.pop(0); cols2.pop(0)
        engagement_data_frame_compare.plot('Timestamp', cols1, ax=axs1, legend=False)
        difficulty_data_frame_compare.plot('Timestamp', cols2, ax=axs2, legend=False)
        fig.text(0.23,0.94,"Number of Samples =" + str(samplenum))
        fig.text(0.65,0.94,"Number of Samples =" + str(samplenum))
        fig.text(0.245,0.90,"Average SSE =" + str(average_sse_engagement))
        fig.text(0.665,0.90,"Average SSE =" + str(average_sse_difficulty))
        # plt.show()
        fig.savefig("Figure_" + str(samplenum) + ".png")
        plt.close(fig)
    print(SSE_dataframe)
    fig, (axs1, axs2) = plt.subplots(1, 2, figsize=(12.0, 4.0))
    SSE_dataframe.boxplot(by='number_of_samples', column=['net_engagement_sse', 'net_difficulty_sse'])
    # SSE_dataframe.plot.box(x = 'number_of_samples', y = 'net_difficulty_sse', ax=axs2, by='number_of_samples')
    plt.show()
    print("Total time = " + str(time.time() - time_initial))
    fig.savefig("BoxPlot_Final.png")


