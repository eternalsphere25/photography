import pandas as pd
import pyexcel_ods3
import re
from sklearn.neighbors import KNeighborsClassifier
import matplotlib.pyplot as plt
import random


def clean_focal_length_text(input_data):
    search_term = '(?<=: ).*(?= mm)'
    for x in range(len(input_data)):
        if x == 0:
            pass
        else:
            new_text = re.search(search_term, input_data[x][0])
            input_data[x][0] = float(new_text.group(0))

def find_empty_items_ods_raw(input_sheet):
    empty = []
    for x in range(len(input_sheet)):
        row = input_sheet[x]
        if len(row) == 0:
            empty.append(x)
    return empty

def find_nearest_neighbor(input_neighbors_list, input_values_list, input_point):
    #Calculate the distance to each number in the input list
    #print('Distance to each number in the input list:')
    distances = [
        input_neighbors_list[y] - input_point for y in 
        range(len(input_neighbors_list))]
    #print(distances)

    #Find largest negative number
    #print('Largest negative number:')
    num_neg = [
        distances[y] for y in range(len(distances)) if distances[y]<0]
    if bool(num_neg) == True:
        max_neg = max(num_neg)
        max_neg_prime = input_neighbors_list[distances.index(max_neg)]
        #print(str(max_neg) + ' (' + str(max_neg_prime) + ')')

    #Find smallest positive number
    #print('Smallest positive number:')
    num_pos = [
        distances[y] for y in range(len(distances)) if distances[y]>0]
    if bool(num_pos) == True:
        min_pos = min(num_pos)
        min_pos_prime = input_neighbors_list[distances.index(min_pos)]
        #print(str(min_pos) + ' (' + str(min_pos_prime) + ')')

    #If the number is at the edge, the adjacent number automatically wins
    if bool(num_neg) == False:
        nearest_neighbor = min_pos_prime
    elif bool(num_pos) == False:
        nearest_neighbor = max_neg_prime
    else:
        #Compare distances and sort according to distance (smaller wins)
        #If distances are equal, then use the photo count as a tiebreaker
        if abs(max_neg) > abs(min_pos):
            nearest_neighbor = min_pos_prime
            #print('Reassignment: ' + str(min_pos_prime))
        elif abs(max_neg) < abs(min_pos):
            nearest_neighbor = max_neg_prime
            #print('Reassignment: ' + str(max_neg_prime))
        else:
            max_neg_count = input_values_list[
                input_neighbors_list.index(max_neg_prime)]
            min_pos_count = input_values_list[
                input_neighbors_list.index(min_pos_prime)]
            #print(str(max_neg_prime) + ' count: ' + str(max_neg_count))
            #print(str(min_pos_prime) + ' count: ' + str(min_pos_count))
            if max_neg_count < min_pos_count:
                nearest_neighbor = min_pos_prime
                #print('Reassignment: ' + str(min_pos_prime))
            elif max_neg_count > min_pos_count:
                nearest_neighbor = max_neg_prime
                #print('Reassignment: ' + str(max_neg_prime))
            else:
                tie = random.randint(0,1)
                if tie == 0:
                    nearest_neighbor = min_pos_prime
                    #print('Reassignment: ' + str(min_pos_prime))
                else:
                    nearest_neighbor = max_neg_prime
                    #print('Reassignment: ' + str(max_neg_prime))
    
    return nearest_neighbor

def generate_dataframe_from_ods_raw(input_data):
    dataframe = pd.DataFrame(input_data[1:], columns=[input_data[0][0], input_data[0][1]])
    return dataframe

def generate_recommendation_list(input_df, input_num):
    lengths = input_df['Focal Length'].tolist()
    counts = input_df['Count'].tolist()
    focal_length_dict = {k:v for k,v in zip(lengths, counts)}

    while len(focal_length_dict) > input_num:
        #print('Current Dict:')
        #print(focal_length_dict)
        value_low = min(focal_length_dict, key=focal_length_dict.get)
        #print('\nLowest value: ' + str(value_low))
        neighbor_list = list(focal_length_dict.keys())
        count_list = list(focal_length_dict.values())
        closest_neighbor = find_nearest_neighbor(neighbor_list, count_list, value_low)
        #print('Reassignment: ' + str(closest_neighbor))


        focal_length_dict[closest_neighbor] += focal_length_dict[value_low]
        focal_length_dict[value_low] = 0

        focal_length_dict = {k:v for k,v in focal_length_dict.items() if v != 0}
        #print('\nNew Dict:')
        #print(focal_length_dict)

    return focal_length_dict

def sort_dict_by_value(input_dict):
    sorted_dict = dict(sorted(input_dict.items(), key=lambda x:x[1],reverse=True))
    return sorted_dict


#------------------------------------------------------------------------------
# PART 1: Import data from file
#-------------------------------------------------------------------------------

#Set file to read
file_name = input('Input file: ')

#Import data (each page of the ods file is a dictionary item)
data = pyexcel_ods3.get_data(file_name)
sheets = list(data.keys())

#print("\nDisplaying data from '" + str(file_name) + "':")
#for item in range(len(sheets)):
#    print("")
#    print('Sheet #' + str(item+1) + ': ' + sheets[item])
#    print(data[sheets[item]])


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
# PART 4: Focal Length Sort
#-------------------------------------------------------------------------------

"""
print('\nInput focal length data:')
print(df_focal_length)

prime_lengths = [14, 15, 16, 18, 20, 24, 28, 35, 50, 70, 85, 105, 120, 135, 200]

lengths = df_focal_length['Focal Length'].tolist()
counts = df_focal_length['Count'].tolist()
df_sorted = df_focal_length.copy()

print('\nChecking for non-prime focal lengths...')
for x in range(len(lengths)):
    if lengths[x] not in prime_lengths:
        non_prime_length = lengths[x]
        print('\n' + str(non_prime_length) + ': NOT PRIME')

        #print('\nPrime focal lengths:')
        #print(prime_lengths)

        #Find nearest prime focal length
        closest_prime = find_nearest_neighbor(
            prime_lengths, counts, non_prime_length)
        
        #Reassign count values to nearest neightbor     
        print('Reassignment: ' + str(closest_prime))
        non_prime_count = df_sorted.loc[
            df_sorted["Focal Length"] == non_prime_length, "Count"]
        prime_count = df_sorted.loc[
            df_sorted["Focal Length"] == closest_prime, "Count"]
    
        new_value = prime_count.item() + non_prime_count.item()

        old_value_index = df_sorted.loc[
            df_sorted["Focal Length"] == non_prime_length, "Count"].index[0]
        new_value_index = df_sorted.loc[
            df_sorted["Focal Length"] == closest_prime, "Count"].index[0]

        df_sorted.at[new_value_index, 'Count'] = new_value
        df_sorted.at[old_value_index, 'Count'] = 0

#Drop rows with empty values for count
df_sorted.drop(df_sorted.loc[df_sorted["Count"] == 0].index, inplace=True)
df_sorted = df_sorted.reset_index(drop=True)

print('\nOriginal focal lengths:')
print(df_focal_length)
print('\nSorted focal lengths:')
print(df_sorted)

#Add in focal length text (mm in 35 mm etc)

focal_length = df_sorted['Focal Length'].tolist()
count = df_sorted['Count'].tolist()

for x in range(len(focal_length)):
    focal_length[x] = str(focal_length[x]) + " mm (35 mm equivalent: " + str(focal_length[x]) + " mm)"

data_out = list(zip(focal_length, count))
df_final = pd.DataFrame(data_out, columns=["Focal Length", "Count"])
print('\nFinal output dataframe:')
print(df_final)
"""

#-------------------------------------------------------------------------------
# PART 5: Most Common Focal Lengths
#-------------------------------------------------------------------------------

#"""
num_recommend = int(
    input('Set number of desired focal length recommendations: '))

print('\nInput focal length data:')
print(df_focal_length)

#Approach #1: Modified nearest neighbor
recommendations = generate_recommendation_list(df_focal_length, num_recommend)
recommendations_sorted = sort_dict_by_value(recommendations)
count = 1
print('\nTop ' + str(num_recommend) + 
' focal length recommendations based on prior usage:')
for key in recommendations_sorted:
    print('#' + str(count) + ': ' + str(key) 
    + ' (' + str(recommendations_sorted[key]) + ')')
    count +=1
#"""

#Approach #2: Logistic regression
quantile = 0.1
quantile_value = df_focal_length.loc[
    df_focal_length["Count"] >= df_focal_length['Count'].quantile(q=quantile)]
while len(quantile_value) > num_recommend:
    quantile += 0.05
    quantile_value = df_focal_length.loc[
        df_focal_length["Count"] >= df_focal_length['Count'].quantile(
            q=quantile)]

print('\nThe following ' + str(num_recommend) + ' lenses are used more than ' + 
"{:.2f}".format(quantile*100) + "% of the time:")
quantile_results = df_focal_length.loc[
    df_focal_length["Count"] >= df_focal_length['Count'].quantile(q=quantile)]
quantile_results_sorted = quantile_results.sort_values(
    by='Count', ascending=False)
print(quantile_results_sorted.to_string(index=False))


"""
Categories:
Ultrawide < 24
Wide 24 - 35
Standard 36 - 69
Short Telephoto 70 - 135
Medium Telephoto 136 - 300
Super Telephoto > 300
"""

# Next task: sort lenses into the above categories
# The next goal is to classify each data point as a certain focal length and 
# then proceed to make two recommendations: (1. Most used per class) (2. Most
# used overall)

# Further out: extract timestamps from EXIF and use to tweak recommendation
# list (lenses used more recently should have more weight)(try using the 
# weighted arithmetic mean equation)





"""
quantile = float(input('\nSet quantile: '))/100
print('\nLenses used more than ' + str(quantile*100) + "% of the time:")
quantile_results = df_focal_length.loc[df_focal_length["Count"] > df_focal_length['Count'].quantile(q=quantile)]
quantile_results_sorted = quantile_results.sort_values(
    by='Count', ascending=False)
print(quantile_results_sorted.to_string(index=False))
"""

#-------------------------------------------------------------------------------
# PART 6: Output to ods
#-------------------------------------------------------------------------------

"""
with pd.ExcelWriter('sorted_focal_lengths.ods', engine='odf') as writer:
    df_final.to_excel(writer, index=False)
"""

"""
#Output scatter plot
focal_lengths = df_focal_length['Focal Length'].tolist()
counts = df_focal_length['Count'].tolist()
fig, ax = plt.subplots()
ax.scatter(focal_lengths, counts)
plt.show()
"""