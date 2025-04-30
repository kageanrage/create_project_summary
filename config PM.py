import re
from pathlib import Path

class Config:
    def __init__(self):
        self.cwd = r'C:\Github local repos\create_project_summary'

        self.projects_dir_path = r'C:\Users\AdeleBascar\OneDrive - STUDENT EDGE\Documents - SE\YouthPanel\SES\Projects'

        self.download_link_example = "https://data.studentedge.org/admin/websiteinsights/export?segmentId="
        self.link_prefix = "https://data.studentedge.org"

        self.downloads_dir = r'C:\Users\AdeleBascar\Downloads'

        self.live_excel_filename = 'Surveys (from 30-08-19)'
        self.live_excel_file_path = r"C:\Users\AdeleBascar\OneDrive - STUDENT EDGE\Documents - SE\YouthPanel"

        self.survey_admin_url = r"https://data.studentedge.org/admin/survey/index"
        self.current_results_dl_url = r"https://data.studentedge.org/admin/survey/download/3f75a648-40a4-4712-9d97-aaf8005ea9a8"
        self.current_results_dl_url_prefix = r"https://data.studentedge.org/admin/survey/download?id="

        self.survey_link_prefix = r'http://surveys.studentedge.com.au/survey/launch?surveyId='
        self.survey_link_middle = '&uid='

        self.sends_dir_template = r'C:\Users\kevin\OneDrive - STUDENT EDGE\Documents - SE\YouthPanel\SES\Projects\client_name\p_number - survey_name\sends'
        self.proj_dir_template = r'C:\Users\kevin\OneDrive - STUDENT EDGE\Documents - SE\YouthPanel\SES\Projects\client_name\p_number - survey_name'

        self.refresh_token = "1000.0bc3f907033b6daba101b22667e8e896.da0773493fae6573569005c5ba51cf6b"

        self.xls_template_simple_full = r'C:\Github local repos\create_project_summary\private\without_rr_data\SP_template.xlsx' # TODO: add to JW and KP desktop
        self.xls_template_complex_full = r'C:\Github local repos\create_project_summary\private\with_rr_data\SP_template.xlsx' # TODO: add to JW and KP desktop

        # self.project_updates_root_dir = r'C:\Users\kevin\OneDrive - STUDENT EDGE\Documents - SE\YouthPanel\SES\project_updates'  # original OneDrive version, as of Mar-20 has sharepoint URL issue
        self.project_updates_root_dir = r'C:\Users\AdeleBascar\Dropbox\project_updates'  # moved to dropbox on 16-Mar-20

        self.json_path = Path('C:\Github local repos\create_project_summary\private\json')