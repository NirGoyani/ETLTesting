from src.utility.ConnectToSharepoint import SharepointConnection
import re
from pathlib import PurePath


class DownloadData:
    def __init__(self):
        self.connection = SharepointConnection()

    def save_file(self, file_name, folder_destination, file_obj):
        file_dir_path = PurePath(folder_destination, file_name)
        with open(file_dir_path, 'wb') as f:
            f.write(file_obj)

    def get_file(self, file_name, folder_name, folder_destination):
        file_obj = self.connection.download_file(file_name, folder_name)
        self.save_file(file_name, folder_destination, file_obj)

    def get_files(self, folder_name, folder_destination):
        files_list = self.connection.get_files_list(folder_name)
        for file in files_list:
            self.get_file(file.name, folder_name, folder_destination)

    def get_latest_file(self, folder_name, folder_destination):
        file_name, file_obj = self.connection.download_latest_file(folder_name)
        self.save_file(file_name, folder_destination, file_obj)

    def get_files_modified_on_latest_date(self, folder_name, folder_destination):
        latest_date = self.connection.get_latest_modified_date(folder_name)
        files_list = self.connection.get_files_modified_on_date(folder_name, latest_date)
        for file in files_list:
            file_obj = self.connection.download_file(file.name, folder_name)
            self.save_file(file.name, folder_destination, file_obj)

    def get_files_by_pattern(self, keyword, folder_name, folder_destination):
        files_list = self.connection.get_files_list(folder_name)
        for file in files_list:
            if re.search(keyword, file.name):
                self.get_file(file.name, folder_name, folder_destination)