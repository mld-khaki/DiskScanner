import os
import pandas as pd
from datetime import datetime
import win32api
import win32con
from PIL import Image
from PIL.ExifTags import TAGS
from mutagen import File as MutagenFile
from tqdm import tqdm

def get_file_attributes(file_path):
    attributes = win32api.GetFileAttributes(file_path)
    attribute_list = []
    if attributes & win32con.FILE_ATTRIBUTE_ARCHIVE:
        attribute_list.append('Archive')
    if attributes & win32con.FILE_ATTRIBUTE_COMPRESSED:
        attribute_list.append('Compressed')
    if attributes & win32con.FILE_ATTRIBUTE_HIDDEN:
        attribute_list.append('Hidden')
    if attributes & win32con.FILE_ATTRIBUTE_READONLY:
        attribute_list.append('Read-Only')
    if attributes & win32con.FILE_ATTRIBUTE_SYSTEM:
        attribute_list.append('System')
    return ', '.join(attribute_list)

def get_exif_data(file_path):
    try:
        image = Image.open(file_path)
        exif_data = image._getexif()
        if exif_data is not None:
            exif = {TAGS.get(tag): value for tag, value in exif_data.items() if tag in TAGS}
            return exif
    except Exception as e:
        return None
    return None

def get_image_info(file_path):
    try:
        image = Image.open(file_path)
        width, height = image.size
        color_depth = len(image.getbands())
        return {'Width': width, 'Height': height, 'Color Depth': color_depth}
    except Exception as e:
        return None

def get_media_info(file_path):
    try:
        audio = MutagenFile(file_path)
        if audio is not None:
            return audio.info.pprint()
    except Exception as e:
        return None
    return None

def get_folder_size(folder_path):
    total_size = 0
    total_size_on_disk = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            total_size += os.path.getsize(file_path)
            try:
                total_size_on_disk += os.path.getsize(file_path)  # For simplicity, we use the same value here
            except OSError:
                continue
    return total_size, total_size_on_disk

def get_text_file_preview(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read(256)
    except Exception as e:
        return None

def scan_folder(source_folder):
    file_data = []
    total_files = sum(len(files) for _, _, files in os.walk(source_folder))
    with tqdm(total=total_files, desc="Scanning Files") as pbar:
        for root, dirs, files in os.walk(source_folder):
            folder_depth = root[len(source_folder):].count(os.sep)
            folder_size, folder_size_on_disk = get_folder_size(root)
            for file in files:
                file_path = os.path.join(root, file)
                file_name = os.path.basename(file_path)
                parent_folder = os.path.dirname(file_path)
                creation_time = datetime.fromtimestamp(os.path.getctime(file_path))
                modification_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                attributes = get_file_attributes(file_path)
                is_compressed = 'Yes' if 'Compressed' in attributes else 'No'
                file_size = os.path.getsize(file_path)
                extension = os.path.splitext(file_path)[1].lower()
                exif_info = get_exif_data(file_path) if extension in ['.jpg', '.jpeg', '.png'] else None
                image_info = get_image_info(file_path) if extension in ['.jpg', '.jpeg', '.png'] and not exif_info else None
                media_info = get_media_info(file_path) if extension in ['.mp3', '.mp4', '.wav'] else None
                text_preview = get_text_file_preview(file_path) if extension in ['.txt', '.log', '.md'] else None

                file_data.append({
                    'File Name': file_name,
                    'Parent Folder': parent_folder,
                    'Folder Depth': folder_depth,
                    'Creation Time': creation_time,
                    'Modification Time': modification_time,
                    'Attributes': attributes,
                    'Compressed': is_compressed,
                    'File Size': file_size,
                    'Extension': extension,
                    'EXIF Data': exif_info,
                    'Image Info': image_info,
                    'Media Info': media_info,
                    'Text Preview': text_preview,
                    'Folder Size': folder_size,
                    'Folder Size on Disk': folder_size_on_disk
                })
                pbar.set_postfix(file=file_name, folder=parent_folder)
                pbar.update(1)
    return file_data

def save_to_excel(data, output_file):
    df = pd.DataFrame(data)
    max_depth = df['Folder Depth'].max() + 1
    level_columns = ['Level ' + str(i) for i in range(1, max_depth + 1)]
    df = df.reindex(columns=[
        'File Name', 'Parent Folder', 'Folder Depth', 'Creation Time', 'Modification Time', 
        'Attributes', 'Compressed', 'File Size', 'Extension', 'EXIF Data', 'Image Info', 
        'Media Info', 'Text Preview', 'Folder Size', 'Folder Size on Disk'
    ] + level_columns)

    for index, row in df.iterrows():
        folder_levels = row['Parent Folder'].split(os.sep)
        for i, level in enumerate(folder_levels):
            df.at[index, 'Level ' + str(i + 1)] = str(level)
        df.at[index, 'Level ' + str(len(folder_levels) + 1)] = str(row['File Name'])
    
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    source_folder = "e://"
    output_file = "D://disk_aaa.xlsx"
    file_data = scan_folder(source_folder)
    save_to_excel(file_data, output_file)
    print(f"File data has been saved to {output_file}")
