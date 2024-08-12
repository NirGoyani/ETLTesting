import environ
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import datetime

# Initialize environment variables
env = environ.Env()
# Read .env file
environ.Env.read_env()

USERNAME = env("SHAREPOINT_EMAIL")
PASSWORD = env("SHAREPOINT_PASSWORD")
SHAREPOINT_SITE = env("SHAREPOINT_URL_SITE")
SHAREPOINT_SITE_NAME = env("SHAREPOINT_SITE_NAME")
SHAREPOINT_DOC = env("SHAREPOINT_DOC_LIBRARY")


class SharepointConnection:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn

    def get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content

    def download_latest_file(self, folder_name):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self.get_files_list(folder_name)
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        file_dict_sorted = {key: value for key, value in sorted(file_dict.items(), key=lambda item: item[1], reverse=True)}
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content

    def get_latest_modified_date(self, folder_name):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self.get_files_list(folder_name)
        latest_date = None
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            if latest_date is None or dt_obj > latest_date:
                latest_date = dt_obj
        return latest_date

    def get_files_modified_on_date(self, folder_name, target_date):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self.get_files_list(folder_name)
        target_date_str = target_date.strftime("%Y-%m-%d")
        filtered_files = []
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            if dt_obj.strftime("%Y-%m-%d") == target_date_str:
                filtered_files.append(file)
        return filtered_files


