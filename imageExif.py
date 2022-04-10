#------------------------------------------------------------------------------
#This program is used for extracting, displaying, and saving photo exif metadata
#The following options are supported:
#   * Single file
#   * Single folder
#   * Single folder tally
#   * Multi-folder (root + subdirectories) tally
#   * Raw EXIF display with ExifTool
#
#The program requires a copy of ExifTool to extract metadata ExifTool can be
#   downloaded for free from its project website:
#   https://exiftool.org/
#
#Note - the program uses a scratch file to send commands to ExifTool. This is
#   to get around ExifTool's inability to work with non-Latin characters. You
#   must manually set the locations of the ExifTool executable and the scratch
#   file (defined below in Part 0).
#
#For metadata tally export, the program saves files in OpenDocument (.ods)
#   format. To open these files, you can use LibreOffice (downloadable for free 
#   from project website):
#   https://www.libreoffice.org/
#
#------------------------------------------------------------------------------

import datetime
import os
import pyexcel_ods3
import re
import subprocess
from decimal import Decimal
from exif import Image
from fractions import Fraction
from imageEXIF_definitions import *
from tqdm import tqdm


class MetadataDict:
    def __init__(self):
        self.metadata = {}
        self.to_ods = []

    def convert_to_list(self, input_dict):
        self.metadata = list(input_dict.items())

    def finalize_data(self, input_dict):
        self.metadata = input_dict

    def format_for_ods(self, input_title):
        self.to_ods = [[input_title, "Count"]] + self.metadata + [""]

    def incr_key(self, input_dict, input_key):
        input_dict[input_key] = input_dict.get(input_key, 0) + 1

    def init_key(self, input_dict, input_key):
        input_dict[input_key] = 1
    
    def insert_null_values(self, expected_total_photos):
        actual_total_photos = sum(self.metadata.values())
        if actual_total_photos != expected_total_photos:
            if no_exif == 0:
                total_diff = expected_total_photos - actual_total_photos
                self.metadata.update({'N/A': total_diff})
            else:
                total_diff = (
                    expected_total_photos - actual_total_photos - no_exif)
                if total_diff != 0:
                    self.metadata.update({'N/A': total_diff})
                self.metadata.update({'No Metadata': no_exif})

    def prepare_for_export(self, input_dict):
        #Finalize the data in the dictionary before export
        self.finalize_data(input_dict)

        #Insert any null values
        if selection == 3:
            expected_total_photos = len(files_list)
        elif selection == 4:
            expected_total_photos = len(full_file_list)
        elif selection == 5:
            expected_total_photos = sub_count
        self.insert_null_values(expected_total_photos)

        #Convert dictionaries into lists
        self.convert_to_list(self.metadata)


def check_if_lens_manual(input_dict):
    if bool(re.match('Manual Lens No CPU', input_dict['LensID'])) == True:
        input_dict['FNumber'] = manual_lens_filler['FNumber']
        input_dict['FocalLength'] = manual_lens_filler['FocalLength']

def check_if_lens_unknown(input_dict):
    if bool(re.match('Unknown', input_dict['LensID'])) == True:
        input_dict['LensID'] = missing_lens_model_metadata[
            input_dict['CameraModelName']]

def convert_to_decimal_degrees(coordinates, reference):
    gps_degrees = coordinates[0]
    gps_minutes = coordinates[1]
    gps_seconds = coordinates[2]
    value = gps_degrees + gps_minutes/60 + gps_seconds/3600
    value = Decimal(value).quantize(Decimal("1.000000"))
    if reference == 'S' or reference == 'W':
        value = -value
    return value

def convert_all_dict_keys_to_str(input_dict):
    keys_values = input_dict.items()
    new_dict = {str(key): value for key, value in keys_values}
    return new_dict

def EXIFTool_data_to_dict(raw_input):
    extracted_exif_list = raw_input.split("\\r\\n")
    output_dict = {}
    for entry in range(len(extracted_exif_list)):
        item = extracted_exif_list[entry].split(":", 1)
        if len(item) == 2:
            item[0] = item[0].replace(" ", "")
            if "b'" in item[0]:
                item[0] = item[0].replace("b'", "")
            item[1] = item[1].replace(" ", "", 1)
            output_dict.update({item[0]: item[1]})  
    return output_dict

def extract_GPS_coords_with_EXIFTool(input_exif_dict, coordinate):
    value = input_exif_dict[coordinate]
    coords_list = []
    coords_list.append(float(re.search("^\d+", value).group(0)))     
    coords_list.append(float(re.search("\d+(?=\\\)", value).group(0)))
    coords_list.append(float(re.search("\d+.\d+(?=\")", value).group(0)))
    coords_list.append(re.search("\w$", value).group(0))
    return coords_list

def extract_metadata_ExifTool(
    program, exiftool_directory, exiftool_writefile, filepath):
    textfile = open(exiftool_writefile, 'w', encoding='utf-8')
    line = filepath
    textfile.write(line)
    textfile.close()

    #Extract metadata from each photo, one by one
    if program == 1:
        filename = re.search("(?<=\/).*", filepath).group(0)
        print('\nExtracting EXIF data from '+ filename +'...')
    exif_data = subprocess.run(
        exiftool_directory + " -charset filename=utf-8 -@ filename_input.txt", 
        capture_output=True)
    extracted_exif = str(exif_data.stdout)
    if program == 1:
        print('Extraction complete')
    return extracted_exif

def find_file_directory(input_file_path):
    dir_raw = input_file_path.split("\\")
    for x in range(2,(len(dir_raw)-1)):
        if x == 2:
            file_path = dir_raw[x]
        else:
            file_path = file_path + "\\" + dir_raw[x]
    return file_path

def format_time(value_time):
	time_string = str(datetime.timedelta(seconds=value_time))
	#Time format is 0:00:000:0000 (hours:minutes:seconds:milliseconds)
	#.split() will separate the format above into its separate chunks
	#The separated chunks will be in a list
	#Hours will be [0], minutes will be [1], etc
	a = time_string.split(':')
	if int(a[0]) != 0:
		print('Session Length:', a[0], 'Hours', a[1], 'Minutes', round(
			float(a[2]), 2), 'Seconds')
	elif int(a[0]) == 0 and int(a[1]) != 0:
		print('Session Length:', a[1], 'Minute(s)', round(
			float(a[2]), 2), 'Second(s)')
	else:
		print('Session Length:', round(float(a[2]), 2), 'Seconds')

def fraction_string_to_decimal(input_string):
    raw_output = input_string.split("/")
    numerator = int(raw_output[0])
    denominator = int(raw_output[1])
    value = numerator/denominator
    return value

def generate_lat_long_coords(input_exif_dict):
    coords_list_lat = extract_GPS_coords_with_EXIFTool(
        input_exif_dict, 'GPSLatitude')
    latitude = convert_to_decimal_degrees(coords_list_lat, coords_list_lat[3])
    coords_list_long = extract_GPS_coords_with_EXIFTool(
        input_exif_dict, 'GPSLongitude')
    longitude = convert_to_decimal_degrees(
        coords_list_long, coords_list_long[3])
    coordinates = [latitude, longitude]
    return coordinates

def get_file_directory(file_path):
    search_regex = re.search(".*(?=\\\)", file_path)
    search_result = search_regex.group(0)
    string = search_result.split("\\")[-1]
    return string

def listEXIFCameraInfoEXIFTool(input_exif_dict):
    print('\nCamera Information:')
    print('- Make: ' + input_exif_dict['Make'])
    print('- Model: ' + input_exif_dict['CameraModelName'])

    #Exclude body serial number if it does not exist
    if 'SerialNumber' in input_exif_dict.keys():
        print('- Body Serial Number: ' + input_exif_dict['SerialNumber'])
    else:
        print('- Body Serial Number: N/A')

    #Show body type (DSLR, point and shoot, phone, etc)
    if input_exif_dict['CameraModelName'] in phones_list:
        print('- Body Type: ' + input_exif_dict['DeviceType'])
    else:
        if input_exif_dict['CameraModelName'] in dslr_list:
            print('- Body Type: Digital SLR')
        elif input_exif_dict['CameraModelName'] in mirrorless_list:
            print('- Body Type: Mirrorless')
        elif input_exif_dict['CameraModelName'] in bridge_list:
            print('- Body Type: Bridge')
        elif input_exif_dict['CameraModelName'] in point_and_shoot_list:
            print('- Body Type: Compact Digital')

def listEXIFGPSInfoEXIFTool(input_exif_dict):
    print('\nGPS Information:')
    coords_lat_long = generate_lat_long_coords(input_exif_dict)
    print('- Latitude: ' + str(coords_lat_long[0]))
    print('- Longitude: ' + str(coords_lat_long[1]))
    print('- Altitude: ' + input_exif_dict['GPSAltitude'])

def listEXIFImageInfoEXIFTool(input_exif_dict):
    print('\nImage Information:')
    #Exposure Mode
    if 'ExposureProgram' in input_exif_dict.keys():
        print('- Exposure Program: ' + input_exif_dict['ExposureProgram'])
    else:
        print('- Exposure Program: N/A')        
    #Shutter, Aperture, ISO, etc
    if 'ExposureTime' in input_exif_dict.keys():
        print('- Exposure Time: ' + input_exif_dict['ExposureTime'])
    else:
        print('- Exposure Time: N/A')
    print('- Aperture: ' + input_exif_dict['FNumber'])
    #ISO
    if 'ISO' in input_exif_dict.keys():
        print('- ISO: ' + input_exif_dict['ISO'])
    else:
        print('- ISO: N/A')
    print('- Focal Length: ' + input_exif_dict['FocalLength'])
    if 'FocalLengthIn35mmFormat' in input_exif_dict.keys():
        print('- Focal Length in 35mm Film: ' + 
        input_exif_dict['FocalLengthIn35mmFormat'])
    else:
        print('- Focal Length in 35mm Film: N/A')

    if 'ExposureCompensation' in input_exif_dict.keys():
        print('- Exposure Compensation: ' + 
        input_exif_dict['ExposureCompensation'])
    else:
        print('- Exposure Compensation: N/A')
    print('- DateTime (Original): ' + input_exif_dict['Date/TimeOriginal'])
    print('- DateTime (Digitized): ' + input_exif_dict['CreateDate'])

def listEXIFLensInfoEXIFTool(input_exif_dict):
    print('\nLens Information:')
    #Lens Manufacturer
    if 'LensMake' in input_exif_dict.keys():
        print('- Lens Manufacturer: ' + input_exif_dict['LensMake'])
    elif input_exif_dict['CameraModelName'] in no_lens_metadata:
        print('- Lens Manufacturer: ' + 
        missing_lens_manufacturer[input_exif_dict['CameraModelName']])
    else:
        print('- Lens Manufacturer: N/A')
    #Lens Name
    if 'LensID' in input_exif_dict.keys():
        print('- Lens Model: ' + input_exif_dict['LensID'])
    elif input_exif_dict['CameraModelName'] in no_lens_metadata:
        print('- Lens Model: ' + 
        missing_lens_model_metadata[input_exif_dict['CameraModelName']])
    else:
        print('- Lens Model: N/A')
    #Lens Serial Number
    if 'LensSerialNumber' in input_exif_dict.keys():
        print('- Lens Serial Number: ' + input_exif_dict['LensSerialNumber'])
    else:        
        print('- Lens Serial Number: N/A')

def missing_metadata_add(input_dict, type):
    if type == 'int':
        value = -1
    elif type == 'float':
        value = -1.0
    elif type == 'str_int':
        value = '-1'
    elif type == 'str_float':
        value = '-1.0'
    return value

def missing_metadata_lens_add(input_dict, missing_info_dict):
    value = missing_info_dict[input_dict['CameraModelName']]
    return value

def missing_metadata_check(input_dict):
    missing_keys = []
    for item in range(len(batch_parameters_EXIF)):
        key = batch_parameters_EXIF[item]
        if key not in input_dict:
            missing_keys.append(key)
    return missing_keys

def print_all_EXIF_dicts():
    print("")
    print('='*80)
    print('\nManufacturers:')
    print_dict_line_by_line(manufacturers.metadata)
    print('\nCameras:')
    print_dict_line_by_line(cameras.metadata)
    print('\nLenses:')
    print_dict_line_by_line(lenses.metadata)
    print('\nShooting Modes:')
    print_dict_line_by_line(mode.metadata)
    print('\nApertures:')
    print_dict_line_by_line(sort_dict_by_key(aperture.metadata, 'float'))
    print('\nShutter Speeds:')
    print_dict_line_by_line(
        sort_dict_by_key(shutter_speed.metadata, 'shutter'))
    print('\nISOs:')
    print_dict_line_by_line(sort_dict_by_key(iso.metadata, 'int'))
    print('\nFocal Lengths:')
    print_dict_line_by_line(
        sort_dict_by_key(focal_length.metadata, 'focal_length'))
    if bool(unclassified.metadata) == True:
        print('\nUnclassified:')
        print_dict_line_by_line(unclassified.metadata)   

def print_dict_line_by_line(input_dict):
    for key, value in input_dict.items():
        print(key, ':', value)

def printListLineByLine(input_list):
    for item in range(len(input_list)):
        print(input_list[item])

def process_metadata(file_to_analyze, extracted_exif_dict):
    #Define EXIF parameters
    batch_parameters_EXIF_mod = batch_parameters_EXIF[:]
    batch_parameters_EXIF_labels_mod = batch_parameters_EXIF_labels[:]

    #Check if there is any missing metadata
    missing_metadata = missing_metadata_check(extracted_exif_dict)
    if len(missing_metadata) != 0:
        process_missing_metadata(
            file_to_analyze, extracted_exif_dict,
            missing_metadata, batch_parameters_EXIF_mod
            )

    #Check if the lens is listed as 'unknown' or is manual
    if 'LensID' in extracted_exif_dict.keys():
        check_if_lens_unknown(extracted_exif_dict)
        check_if_lens_manual(extracted_exif_dict)

    #Convert any 'Hi ISO' values to standard numbers
    if extracted_exif_dict['ISO'] in hi_iso.keys():
        extracted_exif_dict['ISO'] = hi_iso[extracted_exif_dict['ISO']]

    #Convert smartphone focal lengths to 35 mm equivalents:
    if (extracted_exif_dict['CameraModelName'] in phones_list and 
    extracted_exif_dict['FocalLength'] in phone_35mm_conversion):
        extracted_exif_dict['FocalLength'] = phone_35mm_conversion[
            extracted_exif_dict['FocalLength']]

    #Sort EXIF into separate dictionaries (set in definitions file)
    sort_EXIF(extracted_exif_dict, batch_parameters_EXIF_mod)

def process_missing_metadata(
    file_to_analyze, extracted_exif_dict, missing_metadata_list, exif_list):
    #Replace 'ExposureProgram' with 'ExposureMode' if missing
    if 'ExposureProgram' in missing_metadata_list:
        index = exif_list.index('ExposureProgram')
        exif_list[index] = 'ExposureMode'

    #Find the folder name that the file is in
    current_folder = get_file_directory(file_to_analyze)
    if (extracted_exif_dict['CameraModelName'] in 
    missing_metadata_exception_folders.keys()):
        exception_folder_list = missing_metadata_exception_folders[
            extracted_exif_dict['CameraModelName']]
    else:
        exception_folder_list = []

    #Check if the camera has any metadata folder exceptions 
    if (extracted_exif_dict['CameraModelName'] in 
    missing_metadata_exception_folders.keys() and 
    current_folder in exception_folder_list):
        #Check if the current folder matches any of the exception folders
        for item in range(len(exception_folder_list)):
            if current_folder == exception_folder_list[item]:
                #Figure out which exception it is (missing lens, ISO etc),
                # then add the relevant metadata
                if 'ISO' in missing_metadata_list:
                    extracted_exif_dict['ISO'] = missing_metadata_add(
                        extracted_exif_dict, 'int')
                if 'ExposureTime' in missing_metadata_list:
                    extracted_exif_dict['ExposureTime'] = missing_metadata_add(
                        extracted_exif_dict, 'str_float')
                if (extracted_exif_dict['CameraModelName'] in 
                missing_lens_model_metadata.keys() and 
                'LensID' in missing_metadata_list):
                    extracted_exif_dict['LensID'] = missing_metadata_lens_add(
                        extracted_exif_dict, missing_lens_model_metadata)

    else:
        #Add missing metadata
        if 'ISO' in missing_metadata_list:
            extracted_exif_dict['ISO'] = missing_metadata_add(
                extracted_exif_dict, 'int')
        if 'ExposureTime' in missing_metadata_list:
            extracted_exif_dict['ExposureTime'] = missing_metadata_add(
                extracted_exif_dict, 'str_float')
        if (extracted_exif_dict['CameraModelName'] in
         missing_lens_model_metadata.keys() and 
         'LensID' in missing_metadata_list):
            extracted_exif_dict['LensID'] = missing_metadata_lens_add(
                extracted_exif_dict, missing_lens_model_metadata)

def regex_search(search_term, search_text):
    files_list = []
    regex_search = '.' + search_term.upper() + '$'
    regex_search_lowercase = '.' + search_term.lower() + '$'
    for item in range(len(search_text)):
        item_exists = bool(re.search(regex_search, search_text[item]))
        if item_exists == True:
            files_list.append(search_text[item])
        else:
            item_exists = bool(
                re.search(regex_search_lowercase, search_text[item]))
            if item_exists == True:
                files_list.append(search_text[item])
    return files_list

def reverse_dict(input_dict):
    reversed_dict = {v: k for k, v in input_dict.items()}
    return reversed_dict

def sort_dict_by_key(input_dict, mode):
    if mode == 'focal_length':
        #Generate list of focal lengths by extracting relevant numbers
        key_list = input_dict.keys()
        key_list = [x for x in key_list]
        key_list_to_sort = []
        for item in range(len(key_list)):
            key_list_to_sort.append(
                float(re.search('^\d*.\d', key_list[item]).group(0)))
        #Sort the list of extracted numbers
        key_list_to_sort.sort()
        #Convert the list to a sorted dictionary
        sorted_dict = {}
        for item in range(len(key_list_to_sort)):
            for entry in range(len(key_list)):
                #Prevents partial matches (ex 35mm = 135mm)
                term_1 = str(key_list_to_sort[item])
                term_2 = re.search('^\d*.\d', key_list[entry]).group(0)
                if (bool(re.search(str(key_list_to_sort[item]), 
                    str(key_list[entry]))) == True and term_1 == term_2):
                    sorted_dict.update(
                        {key_list[entry]: input_dict[key_list[entry]]})

    elif mode == 'shutter':
        #Generate list of shutter speeds in decimal
        shutter_list = []
        for item in range(len(input_dict.items())):
            key_list = input_dict.keys()
            key_list = [x for x in key_list]
            if bool(re.search("\/", key_list[item])) == True:
                value = fraction_string_to_decimal(key_list[item])
                shutter_list.append(value)
            else:
                value = float(key_list[item])
                shutter_list.append(value)
        #Sort the list
        shutter_list.sort(reverse=True)
        #Convert the sorted values from decimal to fraction
        shutter_list_sorted = []
        for item in range(len(shutter_list)):
            if str(shutter_list[item]) not in shutter_speed_frac_dec:
                shutter_list_sorted.append(
                    str(Fraction(shutter_list[item]).limit_denominator(10000)))
            else:
                shutter_list_sorted.append(
                    shutter_speed_frac_dec[str(shutter_list[item])])
        #Convert the list to a sorted dictionary
        sorted_dict = {}
        for item in range(len(shutter_list_sorted)):
            for entry in range(len(key_list)):
                if shutter_list_sorted[item] == key_list[entry]:
                    sorted_dict.update(
                        {key_list[entry]: input_dict[
                            key_list[entry]]})
                elif (str(shutter_list_sorted[item]) in 
                reverse_dict(shutter_speed_frac_dec)):
                    sorted_dict.update(
                        {shutter_list_sorted[item]: input_dict[
                            key_list[entry]]})

    else:
        #Clean input data by converting all dict keys to string
        input_dict_str = convert_all_dict_keys_to_str(input_dict)
        #Generate list of values
        key_list = input_dict_str.keys()
        if mode == 'int':
            key_list = [int(x) for x in key_list]
        if mode == 'float':
            key_list = [float(x) for x in key_list]
        #Sort list of values
        key_list.sort()
        #Convert the values into a sorted dictionary
        sorted_dict = {}
        for item in range(len(key_list)):
            sorted_dict.update(
                {key_list[item]: input_dict_str[str(key_list[item])]})
    return sorted_dict

def sort_EXIF(input_exif_dict, input_parameters):
    #Extract photo metadata and sort to designated dictionaries
    for entry in range(len(input_parameters)):
        item = input_parameters[entry]
        search_key = input_exif_dict[item]
        if search_key not in tally.metadata:
            tally.init_key(tally.metadata, search_key)
            if item == 'Make':
                manufacturers.init_key(manufacturers.metadata, search_key)
            elif item == 'CameraModelName':
                cameras.init_key(cameras.metadata, search_key)
            elif item == 'LensID':
                lenses.init_key(lenses.metadata, search_key)
            elif item == 'ExposureProgram' or item == 'ExposureMode':
                mode.init_key(mode.metadata, search_key)
            elif item == 'FNumber':
                aperture.init_key(aperture.metadata, search_key)
            elif item == 'ExposureTime':
                shutter_speed.init_key(shutter_speed.metadata, search_key)
            elif item == 'ISO':
                iso.init_key(iso.metadata, search_key)
            elif item == 'FocalLength':
                focal_length.init_key(focal_length.metadata, search_key)
        elif search_key in tally.metadata:
            tally.incr_key(tally.metadata, search_key)
            if item == 'Make':
                manufacturers.incr_key(manufacturers.metadata, search_key)
            elif item == 'CameraModelName':
                cameras.incr_key(cameras.metadata, search_key)
            elif item == 'LensID':
                lenses.incr_key(lenses.metadata, search_key)
            elif item == 'ExposureProgram' or item == 'ExposureMode':
                mode.incr_key(mode.metadata, search_key)              
            elif item == 'FNumber':
                aperture.incr_key(aperture.metadata, search_key)  
            elif item == 'ExposureTime':
                shutter_speed.incr_key(shutter_speed.metadata, search_key)  
            elif item == 'ISO':
                iso.incr_key(iso.metadata, search_key)  
            elif item == 'FocalLength':
                focal_length.incr_key(focal_length.metadata, search_key)

def sort_unclassified_EXIF(metadata_tally_dict):
    metadata_dict_list = [
        manufacturers.metadata,
        cameras.metadata,
        lenses.metadata,
        mode.metadata,
        aperture.metadata,
        shutter_speed.metadata,
        iso.metadata,
        focal_length.metadata
        ]
    # If an entry does not exist in any of the dicts, 
    # add it to the general dict for unclassified items
    for dictionaries in range(len(metadata_dict_list)):
        metadata_tally_dict = ({k: v for k,v in metadata_tally_dict.items() if 
        k not in metadata_dict_list[dictionaries]})
    return metadata_tally_dict


#-------------------------------------------------------------------------------
# PART 0: Define global variables
#-------------------------------------------------------------------------------

#Set path to exiftool.exe and the write file
exiftool_location = "C:/exiftool-12.39/exiftool.exe"
exiftool_write_file = "C:/exiftool-12.39/filename_input.txt"

#Counter for photos with no EXIF data
no_exif = 0

#Initialize metadata objects
tally = MetadataDict()
manufacturers = MetadataDict()
cameras = MetadataDict()
lenses = MetadataDict()
mode = MetadataDict()
aperture = MetadataDict()
shutter_speed = MetadataDict()
iso = MetadataDict()
focal_length = MetadataDict()
unclassified = MetadataDict()

#Set timer start
time_start = datetime.datetime.now()


#-------------------------------------------------------------------------------
# PART 1: Choose mode
#-------------------------------------------------------------------------------

print('\n' + '='*80 + '\n')
print('Establishing connection, please standby...')
print('Connection online')

print('\nOptions: ')
print('1: Single File Mode')
print('2: Single Folder Mode')
print('3: Batch Metadata Checker/Tallier (Single Folder)')
print('4: Batch Metadata Checker/Tallier (Including Subfolders)')
print('5: Batch Metadata Checker/Tallier (Including Subfolders, Separate)')
print('00: Raw Metadata List')
selection = int(input('\nSelection: '))


#-------------------------------------------------------------------------------
# PART 2A: Single file mode
#-------------------------------------------------------------------------------

if selection == 1:

    directory = input('\nInput file directory:')
    input_file = input('Input file name (including file extension): ')
    file_to_analyze = directory + "/" + input_file

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain 
    #non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that
    #are in directories with non-Latin characters)
    extracted_exif = extract_metadata_ExifTool(
        selection, exiftool_location, exiftool_write_file, file_to_analyze)

    #Convert raw output to dictionary
    extracted_exif_dict = EXIFTool_data_to_dict(extracted_exif)

    #Check if image has EXIF, then proceed with EXIF data printout
    print("")
    print('='*80)
    if 'CameraModelName' in extracted_exif_dict.keys():
        print('\nSelected EXIF data:')
        listEXIFCameraInfoEXIFTool(extracted_exif_dict)
        listEXIFLensInfoEXIFTool(extracted_exif_dict)
        listEXIFImageInfoEXIFTool(extracted_exif_dict)
        if (extracted_exif_dict['CameraModelName'] not in 
        dslr_list and 'GPSLatitude' in extracted_exif_dict):
            listEXIFGPSInfoEXIFTool(extracted_exif_dict)
        else:
            print('\nGPS Information:\nNo data available')
    else:
        print('\nImage contains no EXIF data')


#-------------------------------------------------------------------------------
# PART 2B: Folder mode
#-------------------------------------------------------------------------------

elif selection == 2:
    #Set directory and file extension (jpg, png, tiff, etc)
    directory = input('\nInput file directory:')
    file_ext = input('Input image file extension (no dot): ')

    #Find all files of specified type
    print('\nSearching for all ' + file_ext + ' files...')
    files_list_raw = os.listdir(directory)
    files_list = regex_search(file_ext, files_list_raw)

    #List out all data files found in folder
    print('Search complete. ' + str(len(files_list)) + ' file(s) found: ')
    for photo in range(len(files_list)):
        print('\t' + str(photo+1) + '. ' + files_list[photo])
  
    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain 
    #non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that
    #are in directories with non-Latin characters)
    for photo in range(len(files_list)):
        file_to_analyze= str(directory + '/' + files_list[photo] + '\n')
        extracted_exif = extract_metadata_ExifTool(
            selection, exiftool_location, exiftool_write_file, file_to_analyze)

        #Convert raw output to dictionary
        extracted_exif_dict = EXIFTool_data_to_dict(extracted_exif)

        #Check if image has EXIF, then proceed with EXIF data printout
        print("")
        print('='*80)
        print('\nPhoto #' + str(photo+1) + ' of ' + str(len(files_list)) + 
        ': ' + files_list[photo])
        if 'CameraModelName' in extracted_exif_dict.keys():
            print('Selected EXIF data:')
            listEXIFCameraInfoEXIFTool(extracted_exif_dict)
            listEXIFLensInfoEXIFTool(extracted_exif_dict)
            listEXIFImageInfoEXIFTool(extracted_exif_dict)
        else:
            print('No EXIF data, proceeding to next file')


#-------------------------------------------------------------------------------
# PART 2C: Batch metadata tallier
# PART 2D: Batch metadata tallier (including subdirectories)
#-------------------------------------------------------------------------------

elif selection == 3 or selection == 4:
    #Set directory and file extension (jpg, png, tiff, etc)
    directory = input('\nInput file directory:')
    file_ext = input('Input image file extension (no dot): ')

    if selection == 3:
        #Find all files of specified type
        print('\nSearching for all ' + file_ext + ' files...')
        files_list_raw = os.listdir(directory)
        files_list = regex_search(file_ext, files_list_raw) 

        #List out all data files found in folder
        print('Search complete. ' + str(len(files_list)) + ' file(s) found: ')
        for photo in range(len(files_list)):
            print('\t' + str(photo+1) + '. ' + files_list[photo])
  
    elif selection == 4:
        #Find all files of specified type
        print('\nSearching for all files in ' + str(directory) + 
        ' and all subdirectories...')
        root_list = []
        dirs_list = []
        files_list = []
        for root, dirs, files in os.walk(directory):
            root_list.append(root)
            dirs_list.append(dirs)
            files_list.append(files)

        #Generate full path names to each file
        full_file_list_raw = []
        for x in range(len(root_list)):
            for y in range(len(files_list[x])):
                full_file_list_raw.append(
                    os.path.join(root_list[x], files_list[x][y]))
    
        #Include only the files with the specified file extension
        full_file_list = []
        search_upper = file_ext.upper() + ('$')
        search_lower = file_ext.lower() + ('$')
        for item in range(len(full_file_list_raw)):
            if (bool(re.search(search_upper, full_file_list_raw[item])) == True or 
            bool(re.search(search_lower, full_file_list_raw[item])) == True):
                full_file_list.append(full_file_list_raw[item])

        #Display all files in folder and subfolders
        #print(
        # 'Search complete. ' + str(len(full_file_list)) + ' file(s) found: \n')
        #for photo in range(len(full_file_list)):
        #    print('\t' + str(photo+1) + '. ' + full_file_list[photo])

        #Display the number of files found and the number of directories
        print('Search complete. Found ' + str(len(full_file_list)) + ' files in ' + 
        str(len(root_list)-1) + ' subdirectories.')

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain 
    #non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that
    #are in directories with non-Latin characters)
    print('\nExtracting metadata. Be patient, this may take a while...')
    if selection == 3:
        file_list = files_list
    elif selection == 4:
        file_list = full_file_list

    for photo in tqdm(range(len(file_list))):
        if selection == 3:
            file_to_analyze = str(directory + "\\" + file_list[photo] + '\n')
        elif selection == 4:
            file_to_analyze = file_list[photo]

        extracted_exif = extract_metadata_ExifTool(
            selection, exiftool_location, exiftool_write_file, file_to_analyze)

        #Convert raw output to dictionary
        extracted_exif_dict = EXIFTool_data_to_dict(extracted_exif)

        #Precheck #1: Do not process photos with no EXIF data
        if 'CameraModelName' not in extracted_exif_dict.keys():
            no_exif += 1
        else:
            process_metadata(file_to_analyze, extracted_exif_dict)
    print('Metadata extracted successfully')

    #Remove sorted items from original dictionary 
    #to make a list of unclassified items
    print('\nSorting metadata...')
    unclassified.metadata = sort_unclassified_EXIF(tally.metadata)
    print('Sorting complete')

    #Print out results
    print_all_EXIF_dicts()

#-------------------------------------------------------------------------------
# PART 2E: Batch metadata tallier (including subdirectories)(separate folders)
#-------------------------------------------------------------------------------

elif selection == 5:
    #Set directory and file extension (jpg, png, tiff, etc)
    directory = input('\nInput file directory:')
    file_ext = input('Input image file extension (no dot): ')

    #Find all files of specified type
    print('\nSearching for all files in ' + str(directory) + 
    ' and all subdirectories...')
    root_list = []
    dirs_list = []
    files_list = []
    for root, dirs, files in os.walk(directory):
        root_list.append(root)
        dirs_list.append(dirs)
        files_list.append(files)

    #Generate full path names to each file
    full_file_list_raw = []
    for x in range(len(root_list)):
        for y in range(len(files_list[x])):
            full_file_list_raw.append(
                os.path.join(root_list[x], files_list[x][y]))

    #Include only the files with the specified file extension
    full_file_list = []
    search_upper = file_ext.upper() + ('$')
    search_lower = file_ext.lower() + ('$')
    for item in range(len(full_file_list_raw)):
        if (bool(re.search(search_upper, full_file_list_raw[item])) == True or 
        bool(re.search(search_lower, full_file_list_raw[item])) == True):
            full_file_list.append(full_file_list_raw[item])

    #Display all files in folder and subfolders
    #print(
    # 'Search complete. ' + str(len(full_file_list)) + ' file(s) found: \n')
    #for photo in range(len(full_file_list)):
    #    print('\t' + str(photo+1) + '. ' + full_file_list[photo])

    #Display the number of files found and the number of directories
    print('Search complete. Found ' + str(len(full_file_list)) + ' files in ' + 
    str(len(root_list)-1) + ' subdirectories.')

    #Find only the directories that have photos in them
    filled_dirs = []
    for photo in range(len(full_file_list)):
        file_path = find_file_directory(full_file_list[photo])
        if file_path not in filled_dirs:
            filled_dirs.append(file_path)

    print('\nDirectories with photos in them:')
    for item in range(len(filled_dirs)):
        print(str(item+1) + '. ' + filled_dirs[item])

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain 
    #non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that
    #are in directories with non-Latin characters)
    print('\nExtracting metadata. Be patient, this may take a while...')
    file_list = full_file_list

    print('\n' + '='*80)
    print('\nOverall Progress:\n')
    for item in tqdm(range(len(filled_dirs)), desc='Overall'):
        sub_count = 0
        print('\n\nCurrent directory: ' + filled_dirs[item])
        for photo in tqdm(range(len(file_list)), desc='Current', leave=False):
            file_path = find_file_directory(file_list[photo])
            
            if file_path == filled_dirs[item]:
                sub_count += 1
                file_to_analyze = file_list[photo]
                dir_raw = file_list[photo].split('\\')
                dir_current = dir_raw[len(dir_raw)-2]
                extracted_exif = extract_metadata_ExifTool(
                    selection, exiftool_location, exiftool_write_file, file_to_analyze)

                #Convert raw output to dictionary
                extracted_exif_dict = EXIFTool_data_to_dict(extracted_exif)

                #Precheck #1: Do not process photos with no EXIF data
                if 'CameraModelName' not in extracted_exif_dict.keys():
                    no_exif += 1
                else:
                    process_metadata(file_to_analyze, extracted_exif_dict)
        #print('\nMetadata extracted successfully')

        #Remove sorted items from original dictionary 
        #to make a list of unclassified items
        #print('\nSorting metadata...')
        unclassified.metadata = sort_unclassified_EXIF(tally.metadata)
        #print('Sorting complete')

        #Print out results
        #print_all_EXIF_dicts()

        #Prepare EXIF dictionaries for export to disk
        manufacturers.prepare_for_export(
            manufacturers.metadata)
        cameras.prepare_for_export(
            cameras.metadata)
        lenses.prepare_for_export(
            lenses.metadata)
        mode.prepare_for_export(
            mode.metadata)
        aperture.prepare_for_export(
            sort_dict_by_key(aperture.metadata, 'float'))
        shutter_speed.prepare_for_export(
            sort_dict_by_key(shutter_speed.metadata, 'shutter'))
        iso.prepare_for_export(
            sort_dict_by_key(iso.metadata, 'int'))
        focal_length.prepare_for_export(
            sort_dict_by_key(focal_length.metadata, 'focal_length'))
        unclassified.prepare_for_export(
            unclassified.metadata)

        #Combine each list into a few large lists (the library requires each
        #entire page to be a single large list)
        manufacturers.format_for_ods("Manufacturer")
        cameras.format_for_ods("Cameras")
        lenses.format_for_ods("Lenses")
        mode.format_for_ods("Shooting Modes")
        aperture.format_for_ods("Apertures")
        shutter_speed.format_for_ods("Shutter Speed")
        iso.format_for_ods("ISO")
        focal_length.format_for_ods("Focal Length")
        unclassified.format_for_ods("Unclassified")

        write_camera = manufacturers.to_ods + cameras.to_ods + lenses.to_ods
        write_image = (
            mode.to_ods + aperture.to_ods + shutter_speed.to_ods + 
            iso.to_ods + focal_length.to_ods)

        #Write data to disk
        write_out = {"Camera": write_camera, "Image": write_image}
        filename_out = (
            "statistics_" + filled_dirs[item].replace("\\", "_") + ".ods")
        pyexcel_ods3.save_data(filename_out, write_out) 
        #print('\nFile "' + filename_out + '" saved to: ' + os.getcwd())

        #Reset detadata objects for next run
        tally = MetadataDict()
        manufacturers = MetadataDict()
        cameras = MetadataDict()
        lenses = MetadataDict()
        mode = MetadataDict()
        aperture = MetadataDict()
        shutter_speed = MetadataDict()
        iso = MetadataDict()
        focal_length = MetadataDict()
        unclassified = MetadataDict()

        print('\n' + '='*80)
        print('\nOverall Progress:\n')


#-------------------------------------------------------------------------------
# PART 2X: Raw metadata output with ExifTool
#-------------------------------------------------------------------------------

if selection == 00:
    directory = input('\nInput file directory:')
    input_file = input('Input file name (including file extension): ')
    file_to_analyze = directory + "/" + input_file

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain 
    #non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that
    #are in directories with non-Latin characters)
    extracted_exif = extract_metadata_ExifTool(
        selection, exiftool_location, exiftool_write_file, file_to_analyze)

    #Convert raw output to dictionary
    extracted_exif_dict = EXIFTool_data_to_dict(extracted_exif)

    print('\nExtracted EXIF data:')
    print_dict_line_by_line(extracted_exif_dict)


#-------------------------------------------------------------------------------
# PART 3: Output results to file
#-------------------------------------------------------------------------------

if selection == 3 or selection == 4:
    #Prepare EXIF dictionaries for export to disk
    manufacturers.prepare_for_export(
        manufacturers.metadata)
    cameras.prepare_for_export(
        cameras.metadata)
    lenses.prepare_for_export(
        lenses.metadata)
    mode.prepare_for_export(
        mode.metadata)
    aperture.prepare_for_export(
        sort_dict_by_key(aperture.metadata, 'float'))
    shutter_speed.prepare_for_export(
        sort_dict_by_key(shutter_speed.metadata, 'shutter'))
    iso.prepare_for_export(
        sort_dict_by_key(iso.metadata, 'int'))
    focal_length.prepare_for_export(
        sort_dict_by_key(focal_length.metadata, 'focal_length'))
    unclassified.prepare_for_export(
        unclassified.metadata)

    #Combine each list into a few large lists (the library requires each
    #entire page to be a single large list)
    manufacturers.format_for_ods("Manufacturer")
    cameras.format_for_ods("Cameras")
    lenses.format_for_ods("Lenses")
    mode.format_for_ods("Shooting Modes")
    aperture.format_for_ods("Apertures")
    shutter_speed.format_for_ods("Shutter Speed")
    iso.format_for_ods("ISO")
    focal_length.format_for_ods("Focal Length")
    unclassified.format_for_ods("Unclassified")

    write_camera = manufacturers.to_ods + cameras.to_ods + lenses.to_ods
    write_image = (
        mode.to_ods + aperture.to_ods + shutter_speed.to_ods + 
        iso.to_ods + focal_length.to_ods)

    #Write data to disk
    write_out = {"Camera": write_camera, "Image": write_image}
    if selection == 3:
        filename_out = ("statistics_[" + directory.split("\\")[-1] + "]"
        "[single_folder].ods")
    elif selection == 4:
        filename_out = ("statistics_[" + directory.split("\\")[-1] + "]"
        "[with_subfolders].ods")
    pyexcel_ods3.save_data(filename_out, write_out)
    print("")
    print('='*80)
    print('\nFile "' + filename_out + '" saved to: ' + os.getcwd())

#Closing messages
print('\nAll processes completed')
print('Program terminated')

#Display session length
time_end = datetime.datetime.now()
format_time((time_end-time_start).total_seconds())