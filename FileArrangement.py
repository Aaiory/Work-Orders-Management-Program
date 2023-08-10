import os
import re
import shutil
import string

from PIL import Image, ExifTags
from pdf_compressor import compress
from FileInfoFill import write_to_pdf, get_pdf_info
from pypdf import PdfWriter, PdfReader

DEBUGGING = False


def debug_print(*args, **kwargs):
    if DEBUGGING:
        print(*args, **kwargs)


def display_content_in_path(root_folder_path):
    content = [entry.name for entry in os.scandir(root_folder_path)]
    content_str = "\n".join(content)
    debug_print("Content in root folder:\n", content_str)
    return content_str


def get_subfolders(root_directory):
    subfolders = [os.path.join(root_directory, folder) for folder in os.listdir(root_directory) if
                  os.path.isdir(os.path.join(root_directory, folder)) and not any(
                      char in string.ascii_letters for char in folder)]
    return subfolders


def create_folders_and_move_files(root_folder_path):
    for entry in os.scandir(root_folder_path):
        file_path = entry.path
        if entry.is_file() and file_path.endswith(".pdf"):
            file_name = os.path.splitext(entry.name)[0]
            folder_path = os.path.join(root_folder_path, file_name)
            os.makedirs(folder_path, exist_ok=True)
            shutil.move(file_path, os.path.join(folder_path, entry.name))


def copy_pdf_to_subfolders(root_directory, pdf_file):
    subfolders = get_subfolders(root_directory)
    for folder_name in subfolders:
        destination = os.path.join(folder_name, pdf_file)
        if not os.path.exists(destination):
            shutil.copy2(pdf_file, destination)
            debug_print(f"Copied {pdf_file} to {destination}")
        else:
            debug_print(f"{pdf_file} already exists in {destination}")


def fill_pdf_forms(root_directory, pdf_template):
    subfolders = get_subfolders(root_directory)
    for folder_name in subfolders:
        order_num = os.path.basename(folder_name)
        order_filename = os.path.join(folder_name, order_num + ".pdf")
        template_filename = os.path.join(folder_name, pdf_template)
        [location, description] = get_pdf_info(order_filename)
        write_to_pdf(template_filename, order_num, location, description)


def get_work_order_number(pdf):
    for page in pdf.pages:
        content = page.extract_text()
        match = re.search(r'Work Order #\s*(\d+)', content)
        if match:
            return match.group(1)
    return None


def get_new_filename(directory, work_order_number, renamed_numbers):
    existing_files = set(os.listdir(directory))
    base_filename = f"{work_order_number}.pdf"
    i = 1
    while base_filename in existing_files or base_filename in renamed_numbers:
        base_filename = f"{work_order_number} ({i}).pdf"
        i += 1
    return base_filename


def rename_pdf_files(directory):
    renamed_numbers = set()

    for filename in os.listdir(directory):
        if filename.endswith('.pdf') and not filename[:5].isdigit():
            filepath = os.path.join(directory, filename)
            with open(filepath, 'rb') as file:
                pdf = PdfReader(file)
                work_order_number = get_work_order_number(pdf)
                pdf.stream.close()
                if work_order_number:
                    new_filename = get_new_filename(directory, work_order_number, renamed_numbers)
                    renamed_numbers.add(new_filename)
                    new_filepath = os.path.join(directory, new_filename)
                    os.rename(filepath, new_filepath)
                    debug_print(f"Renamed {filename} to {new_filename}")


def create_job_images_folders(root_folder_path):
    for root, dirs, files in os.walk(root_folder_path):
        if root == root_folder_path:
            continue  # Skip the root directory

        job_images_folder = os.path.join(root, "Job Images")

        if not os.path.exists(job_images_folder):
            os.makedirs(job_images_folder)
            debug_print(f"Created '{job_images_folder}' folder.")

        for file in files:
            if any(file.lower().endswith(image_ext) for image_ext in [".jpg", ".jpeg", ".png", ".gif"]):
                source_path = os.path.join(root, file)
                destination_path = os.path.join(job_images_folder, file)
                shutil.copy2(source_path, destination_path)
                debug_print(f"Copied file '{file}' to '{job_images_folder}'.")


def move_files_and_delete_folder(root_folder_path):
    for folder_name, subfolders, filenames in os.walk(root_folder_path):
        if "Job Images" in subfolders:
            job_images_folder = os.path.join(folder_name, "Job Images")
            for file in os.listdir(job_images_folder):
                source = os.path.join(job_images_folder, file)
                destination = os.path.join(folder_name, file)
                shutil.move(source, destination)
            os.rmdir(job_images_folder)


def merge_pdfs_and_images(directory):
    for folder_name in os.listdir(directory):
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            pdf_files = collect_pdf_files(folder_path)
            image_files = collect_image_files(folder_path, pdf_files)
            merge_pdf_and_images(image_files, folder_path)
    debug_print("PDF and image merging completed.")


def collect_pdf_files(folder_path):
    pdf_files = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".pdf"):
            pdf_files.append(os.path.join(folder_path, file_name))
    return pdf_files


def collect_image_files(folder_path, pdf_files):
    image_files = []

    # Add all PDF files to the image_files list
    image_files.extend(pdf_files)

    # Collect all image files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith((".jpg", ".jpeg", ".png")):
            image_files.append(os.path.join(folder_path, filename))

    return image_files


def merge_pdf_and_images(image_files, folder_path):
    output_filename = f"{os.path.basename(os.path.normpath(folder_path))}_merged.pdf"
    output_file = os.path.join(folder_path, output_filename)

    merger = PdfWriter()

    for file in image_files:
        if file.endswith(".pdf"):
            merger.append(file)
        else:
            converted_pdf_path = convert_image_to_pdf(file)
            merger.append(converted_pdf_path)
            os.remove(file)

    merger.write(output_file)
    merger.close()


def convert_image_to_pdf(image_file):
    # Open the image
    image = Image.open(image_file)

    # Rotate the image based on EXIF data
    orientation = next((k for k, v in ExifTags.TAGS.items() if v == 'Orientation'), None)
    if orientation:
        exif = image.getexif()
        if exif and orientation in exif:
            exif_orientation = exif[orientation]
            if exif_orientation == 3:
                image = image.rotate(180, expand=True)
            elif exif_orientation == 6:
                image = image.rotate(270, expand=True)
            elif exif_orientation == 8:
                image = image.rotate(90, expand=True)

    # Calculate the scale factor while maintaining a minimum resolution of 720p
    scale_factor = max(720 / image.width, 720 / image.height)
    new_width, new_height = int(image.width * scale_factor), int(image.height * scale_factor)

    # Resize the image and convert it to RGB
    rgb_image = image.resize((new_width, new_height), Image.ANTIALIAS).convert('RGB')

    # Save the resized and rotated image as a PDF
    pdf_path = os.path.splitext(image_file)[0] + "_converted.pdf"
    rgb_image.save(pdf_path)

    return pdf_path


def delete_converted_pdfs(root_directory):
    for root, dirs, files in os.walk(root_directory):
        for file in files:
            if file.endswith("_converted.pdf"):
                file_path = os.path.join(root, file)
                os.remove(file_path)
                debug_print(f"Deleted file: {file_path}")

    debug_print("Deletion of converted PDF files completed.")


def move_merged_pdfs(root_directory, destination_directory):
    for folder_name, subfolders, filenames in os.walk(root_directory):
        for filename in filenames:
            if filename.endswith("_merged.pdf"):
                source_path = os.path.join(folder_name, filename)
                destination_path = os.path.join(destination_directory, filename)
                shutil.move(source_path, destination_path)
                debug_print(f"Moved file: {filename}")


def compress_pdf_files(directory, power=0):
    for filename in os.listdir(directory):
        if filename.endswith('.pdf'):
            input_file = os.path.join(directory, filename)
            output_file = os.path.join(directory, 'compressed_' + filename)
            compress(input_file, output_file, power)
            os.remove(input_file)


def remove_pre_and_suf(directory):
    for filename in os.listdir(directory):
        new_filename = filename.replace('compressed_', '').replace('_merged', '')
        if new_filename != filename:
            os.rename(os.path.join(directory, filename), os.path.join(directory, new_filename))


def delete_files_and_folders(folder_path: str):
    for entry in os.scandir(folder_path):
        if entry.is_file():
            os.remove(entry.path)
        elif entry.is_dir():
            shutil.rmtree(entry.path)


def copy_paste_fillable_template(folder_path):
    file_name = "Fillable Work order template.pdf"
    src_file = os.path.join(folder_path, file_name)
    if not os.path.exists(src_file):
        shutil.copy(file_name, folder_path)
    else:
        i = 1
        while True:
            new_file_name = f"Fillable Work order template ({i}).pdf"
            new_src_file = os.path.join(folder_path, new_file_name)
            if not os.path.exists(new_src_file):
                shutil.copy(file_name, new_src_file)
                break
            i += 1
