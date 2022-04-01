import math
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyexcel_ods3
import re
from matplotlib.ticker import AutoMinorLocator, MultipleLocator


def bar_graph_basic(input_df, input_labels, ax_labels, title, round_to, orientation, input_file_name, output_file_name):
    #Set text font
    set_default_font()

    #Generate bar graph
    if orientation == 'v':
        fig, ax = plt.subplots(figsize=(10,8))
        ax = input_df['Count'].plot(kind='bar', legend=False)

        #Set axis labels
        ax.set_ylabel(ax_labels[1], fontsize=20, labelpad=12, weight='bold')
        ax.set_xlabel(ax_labels[0], fontsize=20, labelpad=12, weight='bold')

        #Set axis ranges
        max_y_value = find_max_value_in_df_column(input_df, 'Count')
        if round_to == 5:
            ax.set_ylim(0, round_to_nearest_five(max_y_value))
        elif round_to == 10:
            ax.set_ylim(0, round_to_nearest_ten(max_y_value))        
        ax.set_xlim(-1, len(input_df))

        #Set tick numbers and labels
        ax.yaxis.set_major_locator(MultipleLocator(5))
        ax.yaxis.set_minor_locator(AutoMinorLocator(5))
        ax.set_xticklabels(input_labels, rotation=90)

        #Set grid
        ax.grid(which='major', axis='y', linestyle='-', linewidth=0.5)
        ax.set_axisbelow(True)

    elif orientation == 'h':
        fig, ax = plt.subplots(figsize=(8,10))
        ax = input_df['Count'].plot(kind='barh',legend=False)

        #Set axis labels
        ax.set_ylabel(ax_labels[0], fontsize=20, labelpad=12, weight='bold')
        ax.set_xlabel(ax_labels[1], fontsize=20, labelpad=12, weight='bold')

        #Set axis ranges
        max_x_value = find_max_value_in_df_column(input_df, 'Count')
        if round_to == 5:
            ax.set_xlim(0, round_to_nearest_five(max_x_value))
        elif round_to == 10:
            ax.set_xlim(0, round_to_nearest_ten(max_x_value))        
        ax.set_ylim(-1, len(input_df))

        #Set tick numbers and labels
        ax.xaxis.set_major_locator(MultipleLocator(10))
        ax.xaxis.set_minor_locator(AutoMinorLocator(5))
        ax.set_yticklabels(input_labels)

        #Set grid
        ax.grid(which='major', axis='x', linestyle='-', linewidth=0.5)
        ax.set_axisbelow(True)

    #Set tick styles
    ax.tick_params(which='major', length=8, width=2)
    ax.tick_params(which='minor', length=6, width=2)

    #Set graph title
    ax.set_title(title + '\n' + input_file_name, pad=12, fontweight='bold')

    #Generate graph
    plt.savefig(output_file_name, bbox_inches='tight', pad_inches = 0.1)

def clean_focal_length_text(input_data):
    search_term = '(?<=: ).*(?= mm)'
    for x in range(len(input_data)):
        if x == 0:
            pass
        else:
            new_text = re.search(search_term, input_data[x][0])
            input_data[x][0] = float(new_text.group(0))

def dataframe_values_to_list(input_data):
    output_list = []
    for x in range(len(input_data)):
        output_list.append(input_data[x])
    return output_list

def find_empty_items_ods_raw(input_sheet):
    empty = []
    for x in range(len(input_sheet)):
        row = input_sheet[x]
        if len(row) == 0:
            empty.append(x)
    return empty

def find_max_value_in_df_column(input_df, column_name):
    value = input_df[column_name].max()
    return value

def generate_dataframe_from_ods_raw(input_data):
    dataframe = pd.DataFrame(input_data[1:], columns=[input_data[0][0], input_data[0][1]])
    return dataframe

def output_as_bar_graph(input_df, column_name, title_text, xy_axis_labels, round_to, orientation, save_name):
    #Set x axis tick labels
    x_label_values = input_df[column_name].values.tolist()
    x_labels = dataframe_values_to_list(x_label_values)

    #Set axis linewidth
    mpl.rcParams['axes.linewidth'] = 2

    #Generate graph
    bar_graph_basic(input_df, x_labels, xy_axis_labels, title_text, round_to, orientation, file_name, save_name)

def round_to_nearest_five(input_value):
    value = int(str(input_value/10).split('.')[1])
    if value < 6:
        value_whole = int(str(input_value/10).split('.')[0])
        value = (value_whole + 0.5)*10
    else:
        value = math.ceil(input_value/10)*10
    return value

def round_to_nearest_ten(input_value):
    value = math.ceil(input_value/10)*10
    return value

def set_default_font():
    mpl.rcParams['font.size'] = 16
    mpl.rcParams['font.family'] = 'sans-serif'
    mpl.rcParams['font.sans-serif'] = ['Tahoma']
    mpl.rcParams['font.weight'] = 'bold'


#-------------------------------------------------------------------------------
# PART 1: Import data from file
#-------------------------------------------------------------------------------

#Set file to read
file_name = input('Input file: ')

#Import data (each page of the ods file is a dictionary item)
data = pyexcel_ods3.get_data(file_name)
sheets = list(data.keys())

print("\nDisplaying data from '" + str(file_name) + "':")
for item in range(len(sheets)):
    print("")
    print('Sheet #' + str(item+1) + ': ' + sheets[item])
    print(data[sheets[item]])


#-------------------------------------------------------------------------------
# PART 2: Generate dataframes from Sheet 1 (Camera)
#-------------------------------------------------------------------------------

#Set data sheet
sheet = data['Camera']

#Find empty items (empty items separate each section)
empty = find_empty_items_ods_raw(sheet)

#Separate the data on the sheet into separate lists
manufacturers_rows = sheet[0:empty[0]]
cameras_rows = sheet[empty[0]+1:empty[1]]
lenses_rows = sheet[empty[1]+1:empty[2]]

#Assign each data list to a dataframe
df_manufacturers = generate_dataframe_from_ods_raw(manufacturers_rows)
df_cameras = generate_dataframe_from_ods_raw(cameras_rows)
df_lenses = generate_dataframe_from_ods_raw(lenses_rows)


#-------------------------------------------------------------------------------
# PART 3: Generate dataframes from Sheet 2 (Image)
#-------------------------------------------------------------------------------

#Set data sheet
sheet = data['Image']

#Find empty items (empty items separate each section)
empty = find_empty_items_ods_raw(sheet)

#Separate the data on the sheet into separate lists
mode_rows = sheet[0:empty[0]]
aperture_rows = sheet[empty[0]+1:empty[1]]
shutter_speed_rows = sheet[empty[1]+1:empty[2]]
iso_rows = sheet[empty[2]+1:empty[3]]
focal_length_rows = sheet[empty[3]+1:empty[4]]

#Clean focal text length
clean_focal_length_text(focal_length_rows)

#Assign each data list to a dataframe
df_mode = generate_dataframe_from_ods_raw(mode_rows)
df_aperture = generate_dataframe_from_ods_raw(aperture_rows)
df_shutter_speed = generate_dataframe_from_ods_raw(shutter_speed_rows)
df_iso = generate_dataframe_from_ods_raw(iso_rows)
df_focal_length = generate_dataframe_from_ods_raw(focal_length_rows)


#-------------------------------------------------------------------------------
# PART 4: Basic visualizations of data
#-------------------------------------------------------------------------------

#++++++++++++++++++++++++++++++++++++++++
# 4A: Manufacturer
#++++++++++++++++++++++++++++++++++++++++

labels = ['Manufacturer', 'Manufacturer', 'Manufacturer']
output_as_bar_graph(df_manufacturers, 'Manufacturer', labels[0], labels[1:3], 10, 'h', file_name + '_manufacturer.png')

#++++++++++++++++++++++++++++++++++++++++
# 4B: Cameras
#++++++++++++++++++++++++++++++++++++++++

labels = ['Cameras', 'Cameras', 'Count']
output_as_bar_graph(df_cameras, 'Cameras', labels[0], labels[1:3], 10, 'h', file_name + '_cameras.png')

#++++++++++++++++++++++++++++++++++++++++
# 4C: Lenses
#++++++++++++++++++++++++++++++++++++++++

labels = ['Lenses', 'Lenses', 'Count']
output_as_bar_graph(df_lenses, 'Lenses', labels[0], labels[1:3], 10, 'h', file_name + '_lenses.png')

#++++++++++++++++++++++++++++++++++++++++
# 4D: Shooting Mode
#++++++++++++++++++++++++++++++++++++++++

labels = ['Shooting Mode', 'Shooting Mode', 'Count']
output_as_bar_graph(df_mode, 'Shooting Modes', labels[0], labels[1:3], 10, 'h', file_name + '_mode.png')

#++++++++++++++++++++++++++++++++++++++++
# 4E: Aperture
#++++++++++++++++++++++++++++++++++++++++

labels = ['Aperture', 'Aperture', 'Count']
output_as_bar_graph(df_aperture, 'Apertures', labels[0], labels[1:3], 10, 'h', file_name + '_aperture.png')

#++++++++++++++++++++++++++++++++++++++++
# 4F: Shutter Speed
#++++++++++++++++++++++++++++++++++++++++

labels = ['Shutter Speed', 'Shutter Speed', 'Count']
output_as_bar_graph(df_shutter_speed, 'Shutter Speed', labels[0], labels[1:3], 10, 'h', file_name + '_shutter_speed.png')

#++++++++++++++++++++++++++++++++++++++++
# 4G: ISO
#++++++++++++++++++++++++++++++++++++++++

labels = ['ISO', 'ISO', 'Count']
output_as_bar_graph(df_iso, 'ISO', labels[0], labels[1:3], 10, 'h', file_name + '_iso.png')

#++++++++++++++++++++++++++++++++++++++++
# 4H: Focal Length
#++++++++++++++++++++++++++++++++++++++++

labels = ['Focal Length', '35mm Focal Length (mm)', 'Count']
output_as_bar_graph(df_focal_length, 'Focal Length', labels[0], labels[1:3], 10, 'h', file_name + '_focal_length.png')