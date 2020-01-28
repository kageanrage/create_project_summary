import se_admin, se_general
from config import Config
import bs4
import pandas as pd
from pprint import pprint


cfg = Config()

# TODO: scrape survey admin page

# section: scrapes the latest survey admin table
"""
driver, wait = se_general.init_selenium()

se_admin.login_sa(driver, cfg.survey_admin_URL)

soup = se_general.grab_new_soup(driver)

"""

soup_filename = '20200128_soup.txt'

# se_general.export_soup(soup, soup_filename)

dfs = pd.read_html(soup_filename)  # list of dataframes

df = dfs[0]  # the first df in the list

# print(df.info())
# print(df.head())

# TODO: identify which jobs are currently live

live_jobs_df = df[df.Published == True]
# print(live_jobs_df.info())
# print(live_jobs_df.head())

# TODO: OPTIONAL - export df to excel
# live_jobs_df.to_excel('export.xlsx')

# TODO: convert live_jobs_df to a dict of per-project dicts
jobs_reindexed = live_jobs_df.set_index('Project/IO Number')
pprint(jobs_reindexed)

jobs_list = jobs_reindexed.values.tolist()
print(jobs_list)
# TODO: isolate the survey id for each of those, to determine download URL
    # use regex from admin_scrape.py for this
# TODO: add survey_id to dict of dicts
# TODO: download current results for each project
# TODO: create dir structure (todays date = root dir)
# testing phase root dir
root_dir = 'C:\WorkingDir\project_updates'
date_dir_name_template = "YYYY-MM-DD Tue"

# TODO: clone and rename xlsx files from template
fname_template = 'Summary p_num survey_name.xlsx'
# TODO: move downloaded current results files to subdirs
# produce a list of csv files in the downloads dir
# find each one using p_number, then move it to the right dir
# TODO: open all files
# TODO: refresh data on each file (if possible)


