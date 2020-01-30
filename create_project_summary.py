import se_admin, se_general
from config import Config
import bs4, os
import pandas as pd
from pprint import pprint
import datetime
import shutil


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

# section: identify which jobs are currently live

live_jobs_df = df[df.Published == True]
# print(live_jobs_df.info())
# print(live_jobs_df.head())


# section: OPTIONAL - export df to excel
# live_jobs_df.to_excel('export.xlsx')

# section: convert live_jobs_df to a dict of per-project dicts
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

# section: isolate the survey id for each of those, to determine download URL
# section: first, use a simplified version of the regex from admin_scrape.py to grab survey_id and survey_name
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

# section: add survey_id to dict of dicts
for k in live_jobs.keys():
    try:
        live_jobs[k].setdefault('survey_id', all_projects[k]['survey_id'])
    except KeyError:
        print(f"all_projects['{k}'] not found")
        exit()

# pprint(live_jobs)


# section: create (root) dir for todays date
root_dir = 'C:\WorkingDir\project_updates'  # testing phase root dir

todays_date = str(datetime.datetime.today())[0:10].replace('-', '')
today = str(datetime.datetime.today().strftime("%A %d. %B %Y"))[0:3]
date_dir_name = f"{todays_date} {today}"
se_general.create_dir_if_not_exists(f"{root_dir}\\{date_dir_name}")

# section: create dir and SP_files subdir for each project



# section: to use during construction phase only, create one-project dict
short_dict = {}
short_dict['Soft Drink BGS Tracker Jan-20'] = live_jobs['Soft Drink BGS Tracker Jan-20']
pprint(short_dict)


# TODO: create a loop which sequentially does each remaining task to completion, project by project
# TODO: when finished testing, use live_jobs dict for this instead of short_dict

for k in short_dict.keys():
    print(f"Running through loop for short_dict['{k}']")

    # grab key variables from dict
    p_number = live_jobs[k]['p_number']
    client_name = live_jobs[k]['client_name']
    survey_name = live_jobs[k]['survey_name']
    survey_id = live_jobs[k]['survey_id']

    # create directories
    project_dir_name = f"{p_number} {se_general.get_shortened_str(client_name, 5)} {se_general.get_shortened_str(survey_name, 5)}".rstrip()
    project_dir_full = f"{root_dir}\\{date_dir_name}\\{project_dir_name}"
    se_general.create_dir_if_not_exists(project_dir_full)
    project_subdir_full = (f"{root_dir}\\{date_dir_name}\\{project_dir_name}\\SP_files")  # rstrip in case proj_dir_name ends in a space
    se_general.create_dir_if_not_exists(project_subdir_full)

    # TODO: download current results for each project
    dl_link = f"{cfg.download_link_example}{survey_id}"
    # print(dl_link)
    # driver.get(dl_link)  # commented out during test mode

# section: iterate through downloads dir contents to find the downloaded results csv for each job
    most_recent_csv = se_general.identify_cur_res_csv(p_number, cfg.downloads_dir)
    print('most recent csv is:')
    print(most_recent_csv)
    most_recent_csv_full = f"{cfg.downloads_dir}\\{most_recent_csv}"

# section: move downloaded results to appropriate dir, then rename
    shutil.move(most_recent_csv_full, project_subdir_full)  # untested
    moved_csv_full = f"{project_subdir_full}\\{most_recent_csv}"
    desired_csv_name_full = f"{project_subdir_full}\\SurveyResults.csv"
    os.rename(moved_csv_full, desired_csv_name_full)

# section: clone and rename xlsx files from template
    fname_template = f'Summary {p_number} {se_general.get_shortened_str(survey_name, 10)}.xlsx'
    xls_final_filename_full = f"{project_dir_full}\\{fname_template}"
    print(f'attempting to copy {cfg.xls_template_full} to {xls_final_filename_full}\\')
    shutil.copy(cfg.xls_template_full, xls_final_filename_full)

# TODO: open all files
# TODO: refresh data on each file (if possible)
