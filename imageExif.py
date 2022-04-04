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
import warnings
from decimal import Decimal
from exif import Image
from fractions import Fraction
from imageEXIF_definitions import *
from tqdm import tqdm


def checkForEXIF(input_image):
    value = input_image.has_exif
    return value

def check_if_lens_unknown(input_dict):
    if bool(re.match('Unknown', input_dict['LensID'])) == True:
        input_dict['LensID'] = missing_lens_model_metadata[input_dict['CameraModelName']]

def convertToDecimalDegrees(coordinates, reference):
    gps_degrees = coordinates[0]
    gps_minutes = coordinates[1]
    gps_seconds = coordinates[2]
    value = gps_degrees + gps_minutes/60 + gps_seconds/3600
    value = Decimal(value).quantize(Decimal("1.000000"))
    if reference == 'S' or reference == 'W':
        value = -value
    return value

def dict_IncrKey(input_dict, input_key):
    input_dict[input_key] = input_dict.get(input_key, 0) + 1

def dict_InitKey(input_dict, input_key):
    input_dict[input_key] = 1

def EXIFToolDataToDict(raw_input):
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

def extractGPSCoordsWithEXIFTool(input_exif_dict, coordinate):
    value = input_exif_dict[coordinate]
    coords_list = []
    coords_list.append(float(re.search("^\d+", value).group(0)))     
    coords_list.append(float(re.search("\d+(?=\\\)", value).group(0)))
    coords_list.append(float(re.search("\d+.\d+(?=\")", value).group(0)))
    coords_list.append(re.search("\w$", value).group(0))
    return coords_list

def extractMetadataExifTool(program, exiftool_directory, exiftool_writefile, filepath):
    textfile = open(exiftool_writefile, 'w', encoding='utf-8')
    line = filepath
    textfile.write(line)
    textfile.close()

    #Extract metadata from each photo, one by one
    if program == 1:
        filename = re.search("(?<=\/).*", filepath).group(0)
        print('\nExtracting EXIF data from '+ filename +'...')
    exif_data = subprocess.run(exiftool_directory + " -charset filename=utf-8 -@ filename_input.txt", capture_output=True)
    extracted_exif = str(exif_data.stdout)
    if program == 1:
        print('Extraction complete')
    return extracted_exif

def formatTime(value_time):
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

def fractionStringToDecimal(input_string):
    raw_output = input_string.split("/")
    numerator = int(raw_output[0])
    denominator = int(raw_output[1])
    value = numerator/denominator
    return value

def generateLatLongCoords(input_exif_dict):
    coords_list_lat = extractGPSCoordsWithEXIFTool(input_exif_dict, 'GPSLatitude')
    latitude = convertToDecimalDegrees(coords_list_lat, coords_list_lat[3])
    coords_list_long = extractGPSCoordsWithEXIFTool(input_exif_dict, 'GPSLongitude')
    longitude = convertToDecimalDegrees(coords_list_long, coords_list_long[3])
    coordinates = [latitude, longitude]
    return coordinates

def get_file_directory(file_path):
    search_regex = re.search(".*(?=\\\)", file_path)
    search_result = search_regex.group(0)
    string = search_result.split("\\")[-1]
    return string

def listAllEXIFData(input_image):
    for item in range(len(input_image.list_all())):
        print(input_image.list_all()[item])

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
    coords_lat_long = generateLatLongCoords(input_exif_dict)
    print('- Latitude: ' + str(coords_lat_long[0]))
    print('- Longitude: ' + str(coords_lat_long[1]))
    print('- Altitude: ' + input_exif_dict['GPSAltitude'])

def listEXIFImageInfoEXIFTool(input_exif_dict):
    print('\nImage Information:')

    #Exposure Mode
    #if input_exif_dict['CameraModelName'] not in point_and_shoot_list and input_exif_dict['CameraModelName'] != 'Canon EOS DIGITAL REBEL':
    #    print('- Exposure Program: ' + input_exif_dict['ExposureProgram'])
    #else:
    #    print('- Exposure Program: N/A')

    if 'ExposureProgram' in input_exif_dict.keys():
        print('- Exposure Program: ' + input_exif_dict['ExposureProgram'])
    else:
        print('- Exposure Program: N/A')        

    #Shutter, Aperture, ISO, etc
    print('- Exposure Time: ' + input_exif_dict['ExposureTime'])
    print('- Aperture: ' + input_exif_dict['FNumber'])
    print('- ISO: ' + input_exif_dict['ISO'])
    print('- Focal Length: ' + input_exif_dict['FocalLength'])
    if 'FocalLengthIn35mmFormat' in input_exif_dict.keys():
        print('- Focal Length in 35mm Film: ' + input_exif_dict['FocalLengthIn35mmFormat'])
    else:
        print('- Focal Length in 35mm Film: N/A')
    print('- Exposure Compensation: ' + input_exif_dict['ExposureCompensation'])
    print('- DateTime (Original): ' + input_exif_dict['Date/TimeOriginal'])
    print('- DateTime (Digitized): ' + input_exif_dict['CreateDate'])

def listEXIFLensInfoEXIFTool(input_exif_dict):
    print('\nLens Information:')

    #Lens Manufacturer
    if 'LensMake' in input_exif_dict.keys():
        print('- Lens Manufacturer: ' + input_exif_dict['LensMake'])
    elif input_exif_dict['CameraModelName'] in no_lens_metadata:
        print('- Lens Manufacturer: ' + missing_lens_manufacturer[input_exif_dict['CameraModelName']])
    else:
        print('- Lens Manufacturer: N/A')

    #Lens Name
    if 'LensID' in input_exif_dict.keys():
        print('- Lens Model: ' + input_exif_dict['LensID'])
    elif input_exif_dict['CameraModelName'] in no_lens_metadata:
        print('- Lens Model: ' + missing_lens_model_metadata[input_exif_dict['CameraModelName']])
    else:
        print('- Lens Model: N/A')

    #Lens Serial Number
    if 'LensSerialNumber' in input_exif_dict.keys():
        print('- Lens Serial Number: ' + input_exif_dict['LensSerialNumber'])
    else:        
        print('- Lens Serial Number: N/A')

def missing_metadata_add(input_dict, missing_info_dict):
    value = missing_info_dict[input_dict['CameraModelName']]
    return value

def missing_metadata_check(input_dict):
    missing_keys = []
    for item in range(len(batch_parameters_EXIF)):
        key = batch_parameters_EXIF[item]
        if key not in input_dict:
            missing_keys.append(key)
    return missing_keys

def openImageFile():
    directory = input('\nInput file directory:')
    input_file = input('Input file name (including file extension): ')
    file_to_analyze = directory + "/" + input_file
    with open(file_to_analyze, 'rb') as image_file:
        image = Image(image_file)
    return image, file_to_analyze

def orderDictByValue(input_dict):
    sorted_dict = dict(sorted([(v, k) for (k, v) in input_dict.items()], reverse=True))
    ordered_dict = {v: k for (k, v) in sorted_dict.items()}
    return ordered_dict

def printAllEXIFDicts():
    print("")
    print('='*80)
    print('\nManufacturers:')
    printDictLineByLine(manufacturers_EXIF_dict)
    print('\nCameras:')
    printDictLineByLine(cameras_EXIF_dict)
    print('\nLenses:')
    printDictLineByLine(lenses_EXIF_dict)
    print('\nShooting Modes:')
    printDictLineByLine(mode_EXIF_dict)
    print('\nApertures:')
    printDictLineByLine(sortDictByKey(aperture_EXIF_dict, 'float'))
    print('\nShutter Speeds:')
    printDictLineByLine(sortDictByKey(shutter_speed_EXIF_dict, 'shutter'))
    print('\nISOs:')
    printDictLineByLine(sortDictByKey(iso_EXIF_dict, 'int'))
    print('\nFocal Lengths:')
    printDictLineByLine(sortDictByKey(focal_length_EXIF_dict, 'focal_length'))
    if bool(metadata_tally_dict) == True:
        print('\nUnclassified:')
        printDictLineByLine(metadata_tally_dict)   

def printDictLineByLine(input_dict):
    for key, value in input_dict.items():
        print(key, ':', value)

def printListLineByLine(input_list):
    for item in range(len(input_list)):
        print(input_list[item])

def regexSearch(search_term, search_text):
    files_list = []
    regex_search = '.' + search_term.upper() + '$'
    regex_search_lowercase = '.' + search_term.lower() + '$'
    for item in range(len(search_text)):
        item_exists = bool(re.search(regex_search, search_text[item]))
        if item_exists == True:
            files_list.append(search_text[item])
        else:
            item_exists = bool(re.search(regex_search_lowercase, search_text[item]))
            if item_exists == True:
                files_list.append(search_text[item])
    return files_list

def reverseDict(input_dict):
    reversed_dict = {v: k for k, v in input_dict.items()}
    return reversed_dict

def sortDictByKey(input_dict, mode):
    if mode == 'focal_length':
        #Generate list of focal lengths by extracting relevant numbers
        key_list = input_dict.keys()
        key_list = [x for x in key_list]
        key_list_to_sort = []
        for item in range(len(key_list)):
            key_list_to_sort.append(float(re.search('^\d*.\d', key_list[item]).group(0)))
        #Sort the list of extracted numbers
        key_list_to_sort.sort()
        #Convert the list to a sorted dictionary
        sorted_dict = {}
        for item in range(len(key_list_to_sort)):
            for entry in range(len(key_list)):
                #Prevents partial matches (ex 35mm = 135mm)
                term_1 = str(key_list_to_sort[item])
                term_2 = re.search('^\d*.\d', key_list[entry]).group(0)
                if bool(re.search(str(key_list_to_sort[item]), str(key_list[entry]))) == True and term_1 == term_2:
                    sorted_dict.update({key_list[entry]: input_dict[key_list[entry]]})

    elif mode == 'shutter':
        #Generate list of shutter speeds in decimal
        shutter_list = []
        for item in range(len(shutter_speed_EXIF_dict.items())):
            key_list = shutter_speed_EXIF_dict.keys()
            key_list = [x for x in key_list]
            if bool(re.search("\/", key_list[item])) == True:
                value = fractionStringToDecimal(key_list[item])
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
                shutter_list_sorted.append(str(Fraction(shutter_list[item]).limit_denominator(10000)))
            else:
                shutter_list_sorted.append(shutter_speed_frac_dec[str(shutter_list[item])])
        #Convert the list to a sorted dictionary
        sorted_dict = {}
        for item in range(len(shutter_list_sorted)):
            for entry in range(len(key_list)):
                if shutter_list_sorted[item] == key_list[entry]:
                    sorted_dict.update({key_list[entry]: shutter_speed_EXIF_dict[key_list[entry]]})
                elif str(shutter_list_sorted[item]) in reverseDict(shutter_speed_frac_dec):
                    sorted_dict.update({shutter_list_sorted[item]: shutter_speed_EXIF_dict[key_list[entry]]})

    else:
        #Generate list of values
        key_list = input_dict.keys()
        if mode == 'int':
            key_list = [int(x) for x in key_list]
        if mode == 'float':
            key_list = [float(x) for x in key_list]
        #Sort list of values
        key_list.sort()
        #Convert the values into a sorted dictionary
        sorted_dict = {}
        for item in range(len(key_list)):
            sorted_dict.update({key_list[item]: input_dict[str(key_list[item])]})
    return sorted_dict


def sortEXIF(input_exif_dict, metadata_tally_dict, input_parameters):
    #Extract photo metadata and sort to designated dictionaries
    for entry in range(len(input_parameters)):
        item = input_parameters[entry]
        search_key = input_exif_dict[item]
        if search_key not in metadata_tally_dict:
            metadata_tally_dict[search_key] = 1
            if item == 'Make':
                dict_InitKey(manufacturers_EXIF_dict, search_key)
            elif item == 'CameraModelName':
                dict_InitKey(cameras_EXIF_dict, search_key)
            elif item == 'LensID':
                dict_InitKey(lenses_EXIF_dict, search_key)
            elif item == 'ExposureProgram':
                dict_InitKey(mode_EXIF_dict, search_key)
            elif item == 'ExposureMode':
                dict_InitKey(mode_EXIF_dict, search_key)  
            elif item == 'FNumber':
                dict_InitKey(aperture_EXIF_dict, search_key)  
            elif item == 'ExposureTime':
                dict_InitKey(shutter_speed_EXIF_dict, search_key)  
            elif item == 'ISO':
                dict_InitKey(iso_EXIF_dict, search_key)  
            elif item == 'FocalLength':
                dict_InitKey(focal_length_EXIF_dict, search_key)  
        elif search_key in metadata_tally_dict:
            metadata_tally_dict[search_key] = metadata_tally_dict.get(search_key, 0) + 1
            if item == 'Make':
                dict_IncrKey(manufacturers_EXIF_dict, search_key)
            elif item == 'CameraModelName':
                dict_IncrKey(cameras_EXIF_dict, search_key)
            elif item == 'LensID':
                dict_IncrKey(lenses_EXIF_dict, search_key)
            elif item == 'ExposureProgram':
                dict_IncrKey(mode_EXIF_dict, search_key)
            elif item == 'ExposureMode':
                dict_IncrKey(mode_EXIF_dict, search_key)                 
            elif item == 'FNumber':
                dict_IncrKey(aperture_EXIF_dict, search_key)  
            elif item == 'ExposureTime':
                dict_IncrKey(shutter_speed_EXIF_dict, search_key)  
            elif item == 'ISO':
                dict_IncrKey(iso_EXIF_dict, search_key)  
            elif item == 'FocalLength':
                dict_IncrKey(focal_length_EXIF_dict, search_key)

def sortUnclassifiedEXIF(metadata_tally_dict):
    metadata_dict_list = [
        manufacturers_EXIF_dict,
        cameras_EXIF_dict,
        lenses_EXIF_dict,
        mode_EXIF_dict,
        aperture_EXIF_dict,
        shutter_speed_EXIF_dict,
        iso_EXIF_dict,
        focal_length_EXIF_dict
        ]
    #If an entry does not exist in any of the dicts, add it to the general dict for unclassified items
    for dictionaries in range(len(metadata_dict_list)):
        metadata_tally_dict = {k: v for k,v in metadata_tally_dict.items() if k not in metadata_dict_list[dictionaries]}
    return metadata_tally_dict

#Obsolete functions
"""
def listEXIFCameraInfo(input_image):
    print('\nCamera Information:')
    print('- Make: ' + input_image.make)
    print('- Model: ' + input_image.model)
    print('- Body Serial Number: ' + str(input_image.body_serial_number))

def listEXIFGPSInfo(input_image):
    altitude = re.search('[^.][\w]*$', str(input_image.gps_altitude_ref)).group(0)
    latitude = convertToDecimalDegrees(input_image.gps_latitude, input_image.gps_latitude_ref)
    longitude = convertToDecimalDegrees(input_image.gps_longitude, input_image.gps_longitude_ref)
    print('\nGPS Information:')
    print('GPS Version: ' + str(input_image.gps_version_id))
    print('Latitude: ' + str(input_image.gps_latitude) + ' ' + input_image.gps_latitude_ref)
    print('Longitude: ' + str(input_image.gps_longitude) + ' ' + input_image.gps_longitude_ref)
    print('Altitude: ' + str(input_image.gps_altitude) + ' m '+ altitude)
    print('Latitude (Decimal Degrees): ' + str(latitude))
    print('Longtude (Decimal Degrees): ' + str(longitude))

def listEXIFImageInfo(input_image):
    exp_program = re.search('[^.][\w]*$', str(input_image.exposure_program)).group(0)
    print('\nImage Information:')
    print('- Exposure Program: ' + exp_program)
    print('- Exposure Time: ' + str(Fraction(input_image.exposure_time).limit_denominator(100)))
    print('- Aperture: ' + str(input_image.f_number))
    print('- ISO: ' + str(input_image.photographic_sensitivity))
    print('- Focal Length: ' + str(input_image.focal_length) + ' mm')
    print('- Focal Length in 35mm Film: ' + str(input_image.focal_length_in_35mm_film) + 'mm')
    print('- Exposure Compensation: ' + str(input_image.exposure_bias_value))
    print('- DateTime (Original): ' + str(input_image.datetime_original))
    print('- DateTime (Digitized): ' + str(input_image.datetime_digitized))

def listEXIFLensInfo(input_image):
    print('\nLens Information:')
    print('- Lens Manufacturer: ' + input_image.lens_make)
    print('- Lens Model: ' + input_image.lens_model)
    print('- Lens Serial Number: ' + input_image.lens_serial_number)
"""


#-------------------------------------------------------------------------------
# PART 0: Define global variables
#-------------------------------------------------------------------------------

#Ignore the runtime warnings thrown by the module
warnings.filterwarnings("ignore", category=RuntimeWarning)

#Set path to exiftool.exe and the write file
exiftool_location = "C:/exiftool-12.39/exiftool.exe"
exiftool_write_file = "C:/exiftool-12.39/filename_input.txt"

#Counter for photos with no EXIF data
no_exif = 0

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
print('00: Raw Metadata List')
selection = int(input('\nSelection: '))


#-------------------------------------------------------------------------------
# PART 2A: Single file mode
#-------------------------------------------------------------------------------

if selection == 1:

    #Obsolete code - delete later
    """
    #Open file
    image = openImageFile()

    #Check if image has EXIF data before proceeding
    if checkForEXIF(image[0]) == False:
        print('\nERROR: Image has no EXIF data')
        print('Terminating program...')
    elif hasattr(image, 'make') == False:
        print('\nERROR: Image has no EXIF data')
        print('Terminating program...')
    else:
        #Print what kind of EXIF data the file has
       # print('\nList of detected EXIF data:')
        #listAllEXIFData(image)
    """

    directory = input('\nInput file directory:')
    input_file = input('Input file name (including file extension): ')
    file_to_analyze = directory + "/" + input_file

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that are in directories with non-Latin characters)
    extracted_exif = extractMetadataExifTool(selection, exiftool_location, exiftool_write_file, file_to_analyze)

    #Convert raw output to dictionary
    extracted_exif_dict = EXIFToolDataToDict(extracted_exif)

    #print('Extracted EXIF data:')
    #printDictLineByLine(extracted_exif_dict)

    #Check if image has EXIF, then proceed with EXIF data printout
    print("")
    print('='*80)
    if 'CameraModelName' in extracted_exif_dict.keys():
        print('\nSelected EXIF data:')
        listEXIFCameraInfoEXIFTool(extracted_exif_dict)
        listEXIFLensInfoEXIFTool(extracted_exif_dict)
        listEXIFImageInfoEXIFTool(extracted_exif_dict)
        if extracted_exif_dict['CameraModelName'] not in dslr_list and 'GPSLatitude' in extracted_exif_dict:
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
    files_list = regexSearch(file_ext, files_list_raw)

    #List out all data files found in folder
    print('Search complete. ' + str(len(files_list)) + ' file(s) found: ')
    for photo in range(len(files_list)):
        print('\t' + str(photo+1) + '. ' + files_list[photo])
  
    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that are in directories with non-Latin characters)
    for photo in range(len(files_list)):
        file_to_analyze= str(directory + '/' + files_list[photo] + '\n')
        extracted_exif = extractMetadataExifTool(selection, exiftool_location, exiftool_write_file, file_to_analyze)

        #Convert raw output to dictionary
        extracted_exif_dict = EXIFToolDataToDict(extracted_exif)

        #print('Extracted EXIF data:')
        #printDictLineByLine(extracted_exif_dict)

        #Check if image has EXIF, then proceed with EXIF data printout
        print("")
        print('='*80)
        print('\nPhoto #' + str(photo+1) + ' of ' + str(len(files_list)) + ': ' + files_list[photo])
        if 'CameraModelName' in extracted_exif_dict.keys():
            print('Selected EXIF data:')
            listEXIFCameraInfoEXIFTool(extracted_exif_dict)
            listEXIFLensInfoEXIFTool(extracted_exif_dict)
            listEXIFImageInfoEXIFTool(extracted_exif_dict)
        else:
            print('No EXIF data, proceeding to next file')


#-------------------------------------------------------------------------------
# PART 2C: Batch metadata tallier
#-------------------------------------------------------------------------------

elif selection == 3:

    #Set directory and file extension (jpg, png, tiff, etc)
    directory = input('\nInput file directory:')
    file_ext = input('Input image file extension (no dot): ')

    #Find all files of specified type
    print('\nSearching for all ' + file_ext + ' files...')
    files_list_raw = os.listdir(directory)
    files_list = regexSearch(file_ext, files_list_raw)

    #List out all data files found in folder
    print('Search complete. ' + str(len(files_list)) + ' file(s) found: ')
    for photo in range(len(files_list)):
        print('\t' + str(photo+1) + '. ' + files_list[photo])
  
    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that are in directories with non-Latin characters)
    print('\nExtracting metadata. Be patient, this may take a while...')
    initial_loop = False
    for photo in tqdm(range(len(files_list))):
        file_to_analyze = str(directory + "\\" + files_list[photo] + '\n')
        extracted_exif = extractMetadataExifTool(selection, exiftool_location, exiftool_write_file, file_to_analyze)

        #Convert raw output to dictionary
        extracted_exif_dict = EXIFToolDataToDict(extracted_exif)

        #Precheck #1: Do not process photos with no EXIF data
        if 'CameraModelName' not in extracted_exif_dict.keys():
            no_exif += 1
        else:
            #Define EXIF parameters
            batch_parameters_EXIF_mod = batch_parameters_EXIF[:]
            batch_parameters_EXIF_labels_mod = batch_parameters_EXIF_labels[:]

            #Check if there is any missing metadata
            missing_metadata = missing_metadata_check(extracted_exif_dict)
            if len(missing_metadata) != 0:

                print(missing_metadata)

                #Replace 'ExposureProgram' with 'ExposureMode' if missing
                if 'ExposureProgram' in missing_metadata:
                    index = batch_parameters_EXIF_mod.index('ExposureProgram')
                    batch_parameters_EXIF_mod[index] = 'ExposureMode'

                #Find the folder name that the file is in
                current_folder = get_file_directory(file_to_analyze)
                if extracted_exif_dict['CameraModelName'] in missing_metadata_exception_folders.keys():
                    exception_folder_list = missing_metadata_exception_folders[extracted_exif_dict['CameraModelName']]
                else:
                    exception_folder_list = []

                #Check if the camera has any metadata folder exceptions 
                if extracted_exif_dict['CameraModelName'] in missing_metadata_exception_folders.keys() and current_folder in exception_folder_list:
                    #Check if the current folder matches any of the exception folders
                    for item in range(len(exception_folder_list)):
                        if current_folder == exception_folder_list[item]:
                            #Figure out which exception it is (missing lens, ISO etc), then add the relevant metadata
                            if extracted_exif_dict['CameraModelName'] in missing_ISO_metadata.keys() and 'ISO' in missing_metadata:
                                extracted_exif_dict['ISO'] = missing_metadata_add(extracted_exif_dict, missing_ISO_metadata)
                            if extracted_exif_dict['CameraModelName'] in missing_lens_model_metadata.keys() and 'LensID' in missing_metadata:
                                extracted_exif_dict['LensID'] = missing_metadata_add(extracted_exif_dict, missing_lens_model_metadata)

                else:
                    #Add missing metadata
                    if extracted_exif_dict['CameraModelName'] in missing_ISO_metadata.keys() and 'ISO' in missing_metadata:
                        extracted_exif_dict['ISO'] = missing_metadata_add(extracted_exif_dict, missing_ISO_metadata)
                    if extracted_exif_dict['CameraModelName'] in missing_lens_model_metadata.keys() and 'LensID' in missing_metadata:
                        extracted_exif_dict['LensID'] = missing_metadata_add(extracted_exif_dict, missing_lens_model_metadata)  

            #Check if the lens is listed as 'unknown'
            check_if_lens_unknown(extracted_exif_dict)

            #Convert smartphone focal lengths to 35 mm equivalents:
            if extracted_exif_dict['CameraModelName'] in phones_list and extracted_exif_dict['FocalLength'] in phone_35mm_conversion:
                extracted_exif_dict['FocalLength'] = phone_35mm_conversion[extracted_exif_dict['FocalLength']]

            #Sort EXIF into separate dictionaries (set in definitions file)
            sortEXIF(extracted_exif_dict, metadata_tally_dict, batch_parameters_EXIF_mod)
    print('Metadata extracted successfully')

    #print('\nMetadata Tally:')
    #printDictLineByLine(metadata_tally_dict)

    #Remove sorted items from original dictionary to make a list of unclassified items
    print('\nSorting metadata...')
    metadata_tally_dict = sortUnclassifiedEXIF(metadata_tally_dict)
    print('Sorting complete')

    #Print out results
    printAllEXIFDicts()


#-------------------------------------------------------------------------------
# PART 2D: Batch metadata tallier (including subdirectories)
#-------------------------------------------------------------------------------

elif selection == 4:

    #Set directory and file extension (jpg, png, tiff, etc)
    directory = input('\nInput root directory:')
    file_ext = input('Input image file extension (no dot): ')

    #Find all files of specified type
    print('\nSearching for all files in ' + str(directory) + ' and all subdirectories...')
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
            full_file_list_raw.append(os.path.join(root_list[x], files_list[x][y]))
 
    #Include only the files with the specified file extension
    full_file_list = []
    search_upper = file_ext.upper() + ('$')
    search_lower = file_ext.lower() + ('$')
    for item in range(len(full_file_list_raw)):
        if bool(re.search(search_upper, full_file_list_raw[item])) == True or bool(re.search(search_lower, full_file_list_raw[item])) == True:
            full_file_list.append(full_file_list_raw[item])

    #Display all files in folder and subfolders
    #print('Search complete. ' + str(len(full_file_list)) + ' file(s) found: \n')
    #for photo in range(len(full_file_list)):
    #    print('\t' + str(photo+1) + '. ' + full_file_list[photo])

    #Display the number of files found and the number of directories
    print('Search complete. Found ' + str(len(full_file_list)) + ' files in ' + str(len(root_list)-1) + ' subdirectories.')

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that are in directories with non-Latin characters)
    print('\nExtracting metadata. Be patient, this may take a while...')
    initial_loop = False
    for photo in tqdm(range(len(full_file_list))):
        #print(full_file_list[photo])
        file_to_analyze = full_file_list[photo]
        extracted_exif = extractMetadataExifTool(selection, exiftool_location, exiftool_write_file, file_to_analyze)

        #Convert raw output to dictionary
        extracted_exif_dict = EXIFToolDataToDict(extracted_exif)

        #Precheck #1: Do not process phone camera photos (no EXIF)
        if 'CameraModelName' not in extracted_exif_dict.keys():
            no_exif += 1
        else:
            #Define EXIF parameters
            batch_parameters_EXIF_mod = batch_parameters_EXIF[:]
            batch_parameters_EXIF_labels_mod = batch_parameters_EXIF_labels[:]

            #Check if there is any missing metadata
            missing_metadata = missing_metadata_check(extracted_exif_dict)
            if len(missing_metadata) != 0:

                #Replace 'ExposureProgram' with 'ExposureMode' if missing
                if 'ExposureProgram' in missing_metadata:
                    index = batch_parameters_EXIF_mod.index('ExposureProgram')
                    batch_parameters_EXIF_mod[index] = 'ExposureMode'

                #Find the folder name that the file is in
                current_folder = get_file_directory(file_to_analyze)
                if extracted_exif_dict['CameraModelName'] in missing_metadata_exception_folders.keys():
                    exception_folder_list = missing_metadata_exception_folders[extracted_exif_dict['CameraModelName']]
                else:
                    exception_folder_list = []

                #Check if the camera has any metadata folder exceptions 
                if extracted_exif_dict['CameraModelName'] in missing_metadata_exception_folders.keys() and current_folder in exception_folder_list:
                    #Check if the current folder matches any of the exception folders
                    for item in range(len(exception_folder_list)):
                        if current_folder == exception_folder_list[item]:
                            #Figure out which exception it is (missing lens, ISO etc), then add the relevant metadata
                            if extracted_exif_dict['CameraModelName'] in missing_ISO_metadata.keys() and 'ISO' in missing_metadata:
                                extracted_exif_dict['ISO'] = missing_metadata_add(extracted_exif_dict, missing_ISO_metadata)
                            if extracted_exif_dict['CameraModelName'] in missing_lens_model_metadata.keys() and 'LensID' in missing_metadata:
                                extracted_exif_dict['LensID'] = missing_metadata_add(extracted_exif_dict, missing_lens_model_metadata)                    

                else:
                    #Add missing metadata
                    if extracted_exif_dict['CameraModelName'] in missing_ISO_metadata.keys() and 'ISO' in missing_metadata:
                        extracted_exif_dict['ISO'] = missing_metadata_add(extracted_exif_dict, missing_ISO_metadata)
                    if extracted_exif_dict['CameraModelName'] in missing_lens_model_metadata.keys() and 'LensID' in missing_metadata:
                        extracted_exif_dict['LensID'] = missing_metadata_add(extracted_exif_dict, missing_lens_model_metadata)  

            #Check if the lens is listed as 'unknown'
            if 'LensID' in extracted_exif_dict.keys():
                check_if_lens_unknown(extracted_exif_dict)

            #Convert smartphone focal lengths to 35 mm equivalents:
            if extracted_exif_dict['CameraModelName'] in phones_list and extracted_exif_dict['FocalLength'] in phone_35mm_conversion:
                extracted_exif_dict['FocalLength'] = phone_35mm_conversion[extracted_exif_dict['FocalLength']]

            #Sort EXIF into separate dictionaries (set in definitions file)
            sortEXIF(extracted_exif_dict, metadata_tally_dict, batch_parameters_EXIF_mod)
    print('Metadata extracted successfully')

    #Remove sorted items from original dictionary to make a list of unclassified items
    print('\nSorting metadata...')
    metadata_tally_dict = sortUnclassifiedEXIF(metadata_tally_dict)
    print('Sorting complete')

    #Print out results
    printAllEXIFDicts()


#-------------------------------------------------------------------------------
# PART 2X: Raw metadata output with ExifTool
#-------------------------------------------------------------------------------

if selection == 00:
    directory = input('\nInput file directory:')
    input_file = input('Input file name (including file extension): ')
    file_to_analyze = directory + "/" + input_file

    #Run ExifTool
    #WARNING - ExifTool must be on a filepath that does NOT contain non-Latin characters or else it will not run!
    #Write photo file paths to text file (ExifTool cannot read photos that are in directories with non-Latin characters)
    extracted_exif = extractMetadataExifTool(selection, exiftool_location, exiftool_write_file, file_to_analyze)

    #Convert raw output to dictionary
    extracted_exif_dict = EXIFToolDataToDict(extracted_exif)

    print('\nExtracted EXIF data:')
    printDictLineByLine(extracted_exif_dict)


#-------------------------------------------------------------------------------
# PART 3: Output results to file
#-------------------------------------------------------------------------------

if selection == 3 or selection == 4:

    #Prepare EXIF dictionaries for export to disk
    manufacturers = manufacturers_EXIF_dict
    cameras = cameras_EXIF_dict
    lenses = lenses_EXIF_dict
    modes = mode_EXIF_dict
    apertures = sortDictByKey(aperture_EXIF_dict, 'float')
    shutter_speed = sortDictByKey(shutter_speed_EXIF_dict, 'shutter')
    iso = sortDictByKey(iso_EXIF_dict, 'int')
    focal_length = sortDictByKey(focal_length_EXIF_dict, 'focal_length')
    unclassified = metadata_tally_dict

    #Insert any null values
    dict_list = [manufacturers, cameras, lenses, modes, apertures, shutter_speed, iso, focal_length, unclassified]
    if selection == 3:
        expected_total_photos = len(files_list)
    elif selection == 4:
        expected_total_photos = len(full_file_list)
    for item in range(len(dict_list)):
        actual_total_photos = sum(dict_list[item].values())
        if actual_total_photos != expected_total_photos:
            if no_exif == 0:
                total_diff = expected_total_photos - actual_total_photos
                dict_list[item].update({'N/A': total_diff})
            else:
                total_diff = expected_total_photos - actual_total_photos - no_exif
                if total_diff != 0:
                    dict_list[item].update({'N/A': total_diff})
                dict_list[item].update({'No Metadata': no_exif})               

    #Convert dictionaries into lists
    manufacturers = list(manufacturers.items())
    cameras = list(cameras.items())
    lenses = list(lenses.items())
    modes = list(modes.items())
    apertures = list(apertures.items())
    shutter_speed = list(shutter_speed.items())
    iso = list(iso.items())
    focal_length = list(focal_length.items())
    unclassified = list(unclassified.items())

    #Combine each list into a few large lists (the library requires each entire page to be a single large list)
    title_manufacturer = [["Manufacturer", "Count"]] + manufacturers + [""]
    title_cameras = [["Cameras", "Count"]] + cameras + [""]
    title_lenses = [["Lenses", "Count"]] + lenses + [""]
    title_modes = [["Shooting Modes", "Count"]] + modes + [""]
    title_apertures = [["Apertures", "Count"]] + apertures + [""]
    title_shutter_speed = [["Shutter Speed", "Count"]] + shutter_speed + [""]
    title_iso = [["ISO", "Count"]] + iso + [""]
    title_focal_length = [["Focal Length", "Count"]] + focal_length + [""]
    title_unclassified = [['Unclassified', "Count"]] + unclassified + [""]

    write_camera = title_manufacturer + title_cameras + title_lenses
    write_image = title_modes + title_apertures + title_shutter_speed + title_iso + title_focal_length

    #Write data to disk
    write_out = {"Camera": write_camera, "Image": write_image}
    if selection == 3:
        filename_out = "statistics_[" + directory.split("\\")[-1] + "][single_folder].ods"
    elif selection == 4:
        filename_out = "statistics_[" + directory.split("\\")[-1] + "][with_subfolders].ods"
    pyexcel_ods3.save_data(filename_out, write_out)
    print("")
    print('='*80)
    print('\nFile "' + filename_out + '" saved to: ' + os.getcwd())

#Closing messages
print('\nAll processes completed')
print('Program terminated')

#Display session length
time_end = datetime.datetime.now()
formatTime((time_end-time_start).total_seconds())