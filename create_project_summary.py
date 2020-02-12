import se_admin, se_general
from config import Config
import bs4, os, subprocess, datetime, shutil, time, sys, openpyxl
import pandas as pd
from pprint import pprint
import pyautogui
import gui

cfg = Config()


def add_client_name_to_xls():
    wb = openpyxl.load_workbook(str(xls_final_filename_full))
    dash_sheet = wb['Dash']
    dash_sheet['B2'] = client_name
    wb.save(xls_final_filename_full)

# determine if running in auto mode (via BAT file) or not (using GUI for manual inputs)
if len(sys.argv) > 1:
    print('running via script in auto mode - ')
    manually_select_projects = False
    print(f'manually_select_projects = {manually_select_projects}')
    surveys_to_exclude = ['Welcome Survey', 'Education Screener', 'Member Experience Survey']
    print('surveys_to_exclude = ')
    print(surveys_to_exclude)
    manual_inclusions = []
    print('manual_inclusions:')
    print(manual_inclusions)
else:
    manually_select_projects, surveys_to_exclude, manual_inclusions = gui.get_inputs_via_gui()

# scrape the latest survey admin table
driver, wait = se_general.init_selenium()

se_admin.login_sa(driver, cfg.survey_admin_URL)

soup = se_general.grab_new_soup(driver)

soup_filename = '20200128_soup.txt'
se_general.export_soup(soup, soup_filename)

dfs = pd.read_html(soup_filename)  # list of dataframes
all_sa_jobs_df = dfs[0]  # the first df in the list

# identify which jobs are currently live
live_sa_jobs_df = all_sa_jobs_df[all_sa_jobs_df.Published == True]

# OPTIONAL - export df to excel
# live_jobs_df.to_excel('export.xlsx')

# TODO: NB there is a LOT of duplicate code ahead, doing the same stuff for live_jobs as for all_jobs.
#  Re-write, perhaps building 'Live?' variable into all_jobs and then trimming it for use if/when needed
# section: convert live_jobs_df to a dict of per-project dicts
live_sa_jobs_list = live_sa_jobs_df.values.tolist()

live_jobs = {}
for i in range(0, len(live_sa_jobs_list)):
    live_job = {}
    live_job['survey_name'] = live_sa_jobs_list[i][0]
    live_job['p_number'] = live_sa_jobs_list[i][2]
    live_job['client_name'] = live_sa_jobs_list[i][3]
    live_jobs[live_sa_jobs_list[i][0]] = live_job

# in parallel, convert all_sa_jobs_df to a dict of per-project dicts
all_sa_jobs_list = all_sa_jobs_df.values.tolist()

all_jobs = {}
for i in range(0, len(all_sa_jobs_list)):
    job = {}
    job['survey_name'] = all_sa_jobs_list[i][0]
    job['p_number'] = all_sa_jobs_list[i][2]
    job['client_name'] = all_sa_jobs_list[i][3]
    all_jobs[all_sa_jobs_list[i][0]] = job

# isolate the survey id for each of those, to determine download URL.
# First, use a simplified version of the regex from admin_scrape.py to grab survey_id and survey_name
with open(soup_filename) as f:
    canned_soup = bs4.BeautifulSoup(f, "html.parser")

canned_soup_string = str(canned_soup)
mo = cfg.brief_regex.findall(canned_soup_string)
# print(mo[0])

# convert all_projects mo into a dict with p_numbers as keys
all_jobs_ids_and_names = {}
for i in range(0, len(mo)):
    per_project_dict = {}
    per_project_dict['survey_id'] = mo[i][0]
    per_project_dict['survey_name'] = mo[i][1]
    all_jobs_ids_and_names[mo[i][1]] = per_project_dict

# print('All projects dict:')
# pprint(all_projects)

# add survey_id to dict of dicts (Live jobs only)
for k in live_jobs.keys():
    try:
        live_jobs[k].setdefault('survey_id', all_jobs_ids_and_names[k]['survey_id'])
    except KeyError:
        print(f"all_projects['{k}'] not found")
        exit()
print(f'\n\nlive_jobs is of length {len(live_jobs)} and looks like this:')
pprint(live_jobs)

# IN PARALLEL - add survey_id to dict of dicts (all jobs)
for k in all_jobs.keys():
    try:
        all_jobs[k].setdefault('survey_id', all_jobs_ids_and_names[k]['survey_id'])
    except KeyError:
        print(f"all_jobs_ids_and_names['{k}'] not found")
        # TODO: establish why many are not found, then decide if I can accept this. To do with bogus chars like '&'?
        # exit()

print(f'\n\nall_jobs is of length {len(all_jobs)} and looks like this:')
pprint(all_jobs)

# create dir for todays date
root_dir = cfg.project_updates_root_dir
todays_date = str(datetime.datetime.today())[0:10].replace('-', '')
today = str(datetime.datetime.today().strftime("%A %d. %B %Y"))[0:3]
date_dir_name = f"{todays_date} {today}"
se_general.create_dir_if_not_exists(f"{root_dir}\\{date_dir_name}")

# create short one-or-more job dict for manual operation (as opposed to all live jobs)
if manually_select_projects:
    short_dict = {}
    for i in range(0, len(manual_inclusions)):
        short_dict[manual_inclusions[i]] = all_jobs[manual_inclusions[i]]

    print('short_dict looks like this')
    pprint(short_dict)
    jobs_of_interest = short_dict
else:
    jobs_of_interest = live_jobs
    # if show_welcome_survey = False, remove welcome survey from live_jobs before running main loop.
    #  Same for Education and Member FB surveys
    potential_exclusions = ['Welcome Survey', 'Education Screener', 'Member Experience Survey']
    for potential_exclusion in potential_exclusions:
        if potential_exclusion in surveys_to_exclude and jobs_of_interest[f'{potential_exclusion}'] != False:
            del jobs_of_interest[f'{potential_exclusion}']
            f"Deleting {potential_exclusion} from jobs_of_interest"

print('Commencing loop for jobs_of_interest, which looks like this:')
pprint(jobs_of_interest)

for k in jobs_of_interest.keys():
    print(f"Running through loop for jobs_of_interest['{k}']")

    # grab key variables from dict
    p_number = jobs_of_interest[k]['p_number']
    client_name = jobs_of_interest[k]['client_name']
    survey_name = jobs_of_interest[k]['survey_name']
    survey_id = jobs_of_interest[k]['survey_id']

    # create directories
    project_dir_name = f"{p_number} {se_general.get_shortened_str(client_name, 5)} {se_general.get_shortened_str(survey_name, 5)}".rstrip()
    project_dir_full = f"{root_dir}\\{date_dir_name}\\{project_dir_name}"
    project_dir_name, project_dir_full = se_general.create_dir_naming_uniquely_if_exists(project_dir_full)  # adjusts project_dir_name and full if needed
    project_subdir_full = f"{root_dir}\\{date_dir_name}\\{project_dir_name}\\SP_files"  # rstrip in case proj_dir_name ends in a space
    se_general.create_dir_if_not_exists(project_subdir_full)
    print('project_subdir_full is:')
    print(project_subdir_full)

    # Download current results for each project
    dl_link = f"{cfg.current_results_dl_url_prefix}{survey_id}"
    print(dl_link)
    driver.get(dl_link)  # commented out during test mode

    se_general.excel_refresh_all()  # paradoxically placed before opening xls but this should refresh the previous one as late as possible

# iterate through downloads dir contents to find the downloaded results csv for each job
    time.sleep(2)
    most_recent_csv = se_general.loop_verifying_most_recent_csv(p_number, cfg.downloads_dir)
    print('most recent csv is:')
    print(most_recent_csv)
    most_recent_csv_full = f"{cfg.downloads_dir}\\{most_recent_csv}"
    print('most recent csv FULL is:')
    print(most_recent_csv_full)

# move downloaded results to appropriate dir, then rename
    shutil.move(most_recent_csv_full, project_subdir_full)  # untested
    moved_csv_full = f"{project_subdir_full}\\{most_recent_csv}"
    desired_csv_name_full = f"{project_subdir_full}\\SurveyResults.csv"
    os.rename(moved_csv_full, desired_csv_name_full)

# clone and rename xlsx files from template
    fname_template = f'Summary {p_number} {se_general.get_shortened_str(survey_name, 10)}.xlsx'
    xls_final_filename_full = f"{project_dir_full}\\{fname_template}"
    print(f'attempting to copy {cfg.xls_template_full} to {xls_final_filename_full}\\')
    shutil.copy(cfg.xls_template_full, xls_final_filename_full)

    # add_client_name_to_xls()  # commented out as it corrupts the xlsx

# open all files

    subprocess.Popen(f'explorer "{xls_final_filename_full}"')
    time.sleep(10)
    se_general.excel_refresh_all()  # another refresh attempt

# pop up dialog box once script finished
gui.popup_finished_message()

# TODO: figure out how to fully automatically refresh - it seems to be doing it for the second project in the
#  test loop but not the first, and unsure if it's refreshing the last tab

# TODO: maybe close all the xls files and re-open them to force the refresh

# TODO: then merge all 'dashboard/RT' tabs into one xls (noting that I won't be able to refresh after that)

# TODO: add client name to xls template (Dash tab, cell B2) - NB when I try to populate, it corrupts xlsx. Maybe I can try adding it to the CSV, to be pulled through by power query and then used in 'Dash'?