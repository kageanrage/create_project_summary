import se_admin, se_general
from config import Config
import bs4
import pandas as pd
from pprint import pprint

cfg = Config()

# TODO: scrape survey admin page

# section: scrapes the latest survey admin table
'''
driver, wait = se_general.init_selenium()

se_admin.login_sa(driver, cfg.survey_admin_URL)

soup = se_general.grab_new_soup(driver)

'''

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
jobs_list = live_jobs_df.values.tolist()
# print('List of live jobs:')
# print(jobs_list)

# print('First live job:')
# print(jobs_list[0])

live_jobs = {}
for i in range(0, len(jobs_list)):
    live_job = {}
    live_job['survey_name'] = jobs_list[i][0]
    live_job['p_number'] = jobs_list[i][2]
    live_job['client_name'] = jobs_list[i][3]
    live_jobs[jobs_list[i][0]] = live_job

# print('Live jobs dict before adding survey_id')
# pprint(live_jobs)

# TODO: isolate the survey id for each of those, to determine download URL
# TODO: first, use a simplified version of the regex from admin_scrape.py to grab survey_id and survey_name
with open(soup_filename) as f:
    canned_soup = bs4.BeautifulSoup(f, "html.parser")

canned_soup_string = str(canned_soup)
mo = cfg.brief_regex.findall(canned_soup_string)
# print(mo[0])

# convert all_projects mo into a dict with p_numbers as keys
all_projects = {}
for i in range(0, len(mo)):
    per_project_dict = {}
    per_project_dict['survey_id'] = mo[i][0]
    per_project_dict['survey_name'] = mo[i][1]
    all_projects[mo[i][1]] = per_project_dict

# print('All projects dict:')
# pprint(all_projects)

# TODO: add survey_id to dict of dicts
for k in live_jobs.keys():
    try:
        live_jobs[k].setdefault('survey_id', all_projects[k]['survey_id'])
    except KeyError:
        print(f"all_projects['{k}'] not found")
        exit()

# pprint(live_jobs)


# TODO: INCOMPLETE: create dir structure (todays date = root dir)
# testing phase root dir
root_dir = 'C:\WorkingDir\project_updates'
date_dir_name_template = "YYYY-MM-DD Tue"


# TODO: INCOMPLETE: download current results for each project and move them to relevant dir
for v in live_jobs.values():
    print(v['survey_id'])

# TODO: INCOMPLETE: clone and rename xlsx files from template
fname_template = 'Summary p_num survey_name.xlsx'
# TODO: INCOMPLETE: move downloaded current results files to subdirs
# produce a list of csv files in the downloads dir
# find each one using p_number, then move it to the right dir
# TODO: INCOMPLETE: open all files
# TODO: INCOMPLETE: refresh data on each file (if possible)
