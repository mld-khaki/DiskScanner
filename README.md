# DriveScanner

DriveScanner is a Python program that recursively scans a specified directory, collecting detailed information about files and folders. It saves the data in an Excel file, including attributes like file names, creation times, modification times, sizes, and more. It also gathers metadata for media files and previews text content for text files.

## Features

- Recursive directory scanning
- Collects file and folder details: name, parent folder, creation time, modification time, attributes, compression status, file size, extension, EXIF data (for images), media info (for audio/video), and text previews (for text files)
- Computes folder sizes (real size and size on disk)
- Saves data to an Excel file with a tabbed representation of directory structure
- Displays progress during scanning

## Requirements

- Python 3.6+
- pandas
- pywin32
- Pillow
- mutagen
- tqdm

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/drivescanner.git
   cd drivescanner

Install the required libraries:
bash
Copy code
pip install pandas pywin32 Pillow mutagen tqdm
Usage
Update the source_folder and output_file variables in the script to your desired input directory and output Excel file path.
Run the script:
bash


if __name__ == "__main__":
    source_folder = "e://"
    output_file = "D://disk_aaa.xlsx"
    file_data = scan_folder(source_folder)
    save_to_excel(file_data, output_file)
    print(f"File data has been saved to {output_file}")
    
## Functions
### get_file_attributes(file_path)
Returns file attributes (archive, compressed, hidden, read-only, system).

### get_exif_data(file_path)
Returns EXIF data for images.

### get_image_info(file_path)
Returns image dimensions and color depth if EXIF data is unavailable.

### get_media_info(file_path)
Returns media info for audio and video files.

### get_folder_size(folder_path)
Returns the total size and size on disk of a folder.

### get_text_file_preview(file_path)
Returns the first 256 characters of a text file.

### scan_folder(source_folder)
Scans the directory and collects data for all files and folders.

### save_to_excel(data, output_file)
Saves the collected data to an Excel file.

## Contributing
Feel free to fork the repository and submit pull requests. For major changes, please open an issue to discuss your proposed changes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Contact
If you have any questions or suggestions, please open an issue or contact me at first name dot last name @ gmail.com

