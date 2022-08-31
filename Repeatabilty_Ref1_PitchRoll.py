import pandas as pd
import glob
import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import linregress
import statistics

excel_data_path = "C:/Users/Advan/Documents/Precitec/Data/REPEATABILITY/Another_Ivy_Test_Data/Excel_Data/Ref1/*.xlsx"
raw_data_path = "C:/Users/Advan/Documents/Precitec/Data/REPEATABILITY/Another_Ivy_Test_Data/Raw_Data/Ref1/Fast_Recipe_30um_spacing_Thickness 1_0.86699998x0.852mm_29.896551x29.37931µm_8000Hz_8%.asc"

#Retrive pixel size from raw data for x data points
a_file = open(raw_data_path)
file_contents = a_file.read()
contents_split = file_contents.splitlines()
a_file.close()

xpixlesValue = contents_split[2]
ypixlesValue = contents_split[3]
pitchValue = contents_split[4]
rollValue = contents_split[5]

xpixlesValue = xpixlesValue[13:]
ypixlesValue = ypixlesValue[13:]
pitchValue = pitchValue[13:]
rollValue = rollValue[13:]

xPixles = int(xpixlesValue)
yPixels = int(ypixlesValue)
pitch = float(pitchValue)
roll = float(rollValue)

pixel_size_ref1_pitch = float("{:.2f}".format(pitch/xPixles))/1000
pixel_size_ref1_roll = float("{:.2f}".format(roll/yPixels))/1000

#Retrieve excel data
all_data_pitch = []
all_data_roll = []
sa_angle_data_pitch = []
sa_angle_data_roll = []
for f in glob.glob(excel_data_path):
    df = pd.read_excel(f)
    df_shape = df.shape
   
    # define the region of interest
    matrix = df.iloc[0:df_shape[0],0:df_shape[1]] #AOI_Pitch and Roll
       
    # define if pitch or roll measurment
    y_pitch = matrix.mean(axis='index') #pitch along the rows
    y_roll = matrix.mean(axis='columns') #roll along the columns
      
    # matix size
    ref1_pitch_data_points = (df.shape[1]-0)*pixel_size_ref1_pitch
    ref1_roll_data_points = (df.shape[0]-0)*pixel_size_ref1_roll
    
    # change matrix size and pixel size
    x_pitch = np.arange(0,ref1_pitch_data_points,pixel_size_ref1_pitch)
    x_roll = np.arange(0,ref1_roll_data_points,pixel_size_ref1_roll)
    
    #pitchConvert a
    slope_pitch, intercept, r_value, p_value, std_err = linregress(x_pitch,y_pitch)
    sa_angle_pitch = np.arctan(slope_pitch)*(180/np.pi)
    print('pitch angle;', sa_angle_pitch)  
    #roll
    slope_roll, intercept, r_value, p_value, std_err = linregress(x_roll,y_roll)
    sa_angle_roll = np.arctan(slope_roll)*(180/np.pi)
    print('roll angle;',sa_angle_roll) 
    
    # print(sa_angle)
    # plt.title("Line graph")
    # plt.xlabel("X axis")
    # plt.ylabel("Y axis")
    # plt.scatter(x, y, color ="red")
    # plt.show()
    
    all_data_pitch.append(y_pitch)
    all_data_roll.append(y_roll)
    sa_angle_data_pitch.append(sa_angle_pitch)
    sa_angle_data_roll.append(sa_angle_roll)
      
all_data_pitch = pd.concat(all_data_pitch)
all_data_roll = pd.concat(all_data_roll)
all_data_pitch.to_excel('Another_Ivy_Test_Ref1_Pitch.xlsx')  
all_data_roll.to_excel('Another_Ivy_Test_Ref1_Roll.xlsx')  

stdevALL_pitch = statistics.stdev(sa_angle_data_pitch)
stdevALL_roll = statistics.stdev(sa_angle_data_roll)
print("Another_Ivy_Test_Ref1_Pitch;",stdevALL_pitch)
print("Another_Ivy_Test_Ref1_Roll;",stdevALL_roll)