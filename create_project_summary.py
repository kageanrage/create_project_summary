import se_admin, se_general
from config import Config
import bs4
import pandas as pd


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
print(live_jobs_df.info())
print(live_jobs_df.head())

# TODO: isolate the survey id for each of those, to determine download URL
# TODO: download current results for each of those
# TODO: create dir structure (todays date = root dir)
# TODO: clone and rename xlsx files from template
# TODO: move downloaded current results files to subdirs
# TODO: open all files
# TODO: refresh data on each file (if possible)
