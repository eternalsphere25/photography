manufacturers_list = [
    'NIKON CORPORATION'
    ]

dslr_list = [
    'NIKON D80',
    'NIKON D600',
    'Canon EOS DIGITAL REBEL XT',
    'Canon EOS DIGITAL REBEL'
    ]

mirrorless_list = [
    'NIKON Z 6_2'
    ]

phones_list = [
    'SAMSUNG-SGH-I747',
    'SAMSUNG-SM-G900A',
    'SM-G950F'
    ]

point_and_shoot_list = [
    'Canon DIGITAL IXUS 55',
    'Canon PowerShot S45',
    'COOLPIX L4',
    'COOLPIX L10',
    'COOLPIX L12'
    ]

exposure_programs_list = [
    'Not Defined',
    'Aperture-priority AE',
    'Shutter speed priority AE',
    'Manual',
    'Auto'
    ]

lens_list_z_mount = [
    "NIKKOR Z 24-120mm f/4 S"
    ]

lens_list_f_mount_FX = [
    "AF-S Nikkor 35mm f/1.8G ED",
    "AF Nikkor 50mm f/1.8D",
    "AF-S Nikkor 50mm f/1.8G",
    "AF-S Zoom-Nikkor 14-24mm f/2.8G ED"
    "AF-S Nikkor 24-120mm f/4G ED VR",
    "AF-S Nikkor 70-200mm f/4G ED VR",
    "AF-S VR Micro-Nikkor 105mm f/2.8G IF-ED",
    ]

lens_list_f_mount_DX = [
    "AF-S DX Nikkor 35mm f/1.8G",
    "AF-S DX VR Zoom-Nikkor 16-85mm f/3.5-5.6G ED",
    "AF-S DX Zoom-Nikkor 18-135mm f/3.5-5.6G IF-ED",
    "AF-S DX VR Zoom-Nikkor 18-200mm f/3.5-5.6G IF-ED [II]",
    "AF-S DX VR Zoom-Nikkor 18-200mm f/3.5-5.6G IF-ED [II] or AF-S DX VR Zoom-Nikkor 18-200mm f/3.5-5.6G IF-ED",
    "AF-S DX VR Zoom-Nikkor 18-200mm f/3.5-5.6G IF-ED or AF-S DX VR Zoom-Nikkor 18-200mm f/3.5-5.6G IF-ED [II]",
    ]

iso_list = [
    '100',
    '125',
    '160',
    '200',
    '250',
    '320',
    '400',
    '500',
    '640',
    '800',
    '1000',
    '1250',
    '1600',
    '2000',
    '2500',
    '3200',
    '4000',
    '5000',
    '6400',
    '8000',
    '10000',
    '12800',
    '16000',
    '20000',
    '25600',
    '32000',
    '40000',
    '51200'
    ]

aperture_list = [
    '1.8',
    '2',
    '2.2',
    '2.5',
    '2.8',
    '3.2',
    '3.5',
    '4',
    '4.5',
    '5',
    '5.6',
    '6.3',
    '7.1',
    '8',
    '9',
    '10',
    '11',
    '13',
    '14',
    '16',
    '18',
    '20',
    '22'
    ]

batch_parameters_EXIF = [
    'Make',
    'CameraModelName',
    'LensID',
    'ExposureProgram',
    'FNumber',
    'ExposureTime',
    'ISO',
    'FocalLength',
    ]

batch_parameters_EXIF_labels = [
    'Manufacturers',
    'Cameras',
    'Lenses',
    'Shooting Modes',
    'Apertures',
    'Shutter Speeds',
    'ISOs',
    'Focal Lengths'
    ]

shutter_speed_frac_dec = {
    '0.6': '1/1.6',
    '0.5': '1/2',
    '0.4': '1/2.5', 
    '0.3': '1/3'
    }

metadata_tally_dict = {}
manufacturers_EXIF_dict = {}
cameras_EXIF_dict = {}
lenses_EXIF_dict = {}
mode_EXIF_dict = {}
aperture_EXIF_dict = {}
shutter_speed_EXIF_dict = {}
iso_EXIF_dict = {}
focal_length_EXIF_dict = {}

no_exposure_program = [
    'Canon PowerShot S45',
    'Canon EOS DIGITAL REBEL',
    'Canon DIGITAL IXUS 55'
    ]

no_lens_metadata = [
    'Canon EOS DIGITAL REBEL XT',
    'Canon EOS DIGITAL REBEL',
    'SAMSUNG-SGH-I747',
    'SAMSUNG-SM-G900A',
    'SM-G950F',
    'Canon DIGITAL IXUS 55',
    'Canon PowerShot S45',
    'COOLPIX L4',
    'COOLPIX L10',
    'COOLPIX L12'
    ]

missing_lens_model_metadata = {
    'SAMSUNG-SGH-I747': 'Samsung 3.7 mm (35 mm equivalent: 26.68 mm); f/2.6',
    'SAMSUNG-SM-G900A': 'Samsung 4.8 mm (35 mm equivalent: 31 mm); f/2.2',
    'SM-G950F': 'Samsung 4.2 mm (35 mm equivalent: 26 mm); f/1.7',
    'Canon DIGITAL IXUS 55': 'Canon 5.8-17.4 mm (35 mm equivalent: 35-105 mm); f/2.8-4.9',
    'Canon PowerShot S45': 'Canon 7.1-21.3 mm zoom lens (35 mm equivalent: 35-105 mm); f/2.8-4.9',
    'COOLPIX L4': 'Nikkor 6.3-18.9 mm (35 mm equivalent: 38-114 mm); f/2.8-4.9',
    'COOLPIX L10': 'Nikkor 6.2-18.6 mm (35 mm equivalent: 37.5-112.5 mm); f/2.8-5.2',
    'COOLPIX L12': 'Nikkor 5.7-17.1 mm (35 mm equivalent: 35-105 mm); f/2.8-4.7'
    }

missing_lens_manufacturer = {
    'SAMSUNG-SGH-I747': 'Samsung',
    'SAMSUNG-SM-G900A': 'Samsung',
    'SM-G950F': 'Samsung',
    'Canon DIGITAL IXUS 55': 'Canon',
    'Canon PowerShot S45': 'Canon',
    'COOLPIX L4': 'Nikon',
    'COOLPIX L10': 'Nikon',
    'COOLPIX L12': 'Nikon'
    }

phone_35mm_conversion = {
    '3.7 mm': '3.7 mm (35 mm equivalent: 26.68 mm)',
    '4.8 mm': '4.8 mm (35 mm equivalent: 31 mm)',
    '4.2 mm': '4.2 mm (35 mm equivalent: 26 mm)'
    }