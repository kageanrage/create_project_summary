import se_admin, se_general
from config import Config
import bs4, os, subprocess, datetime, shutil, time, sys, openpyxl
import pandas as pd
from pprint import pprint
import pyautogui
import gui
import win32com.client as win32


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
    # close_excel = True  # turned off 11-05-20 as feature doesnt work
    # print(f'close_excel = {close_excel}')  # turned off 11-05-20 as feature doesnt work
else:
    manually_select_projects, surveys_to_exclude, manual_inclusions = gui.get_inputs_via_gui()

# scrape the latest survey admin table

driver, wait = se_general.init_selenium()
live_jobs, all_jobs = se_admin.grab_sa_projects_dicts(driver)


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

# excel = win32.gencache.EnsureDispatch('Excel.Application')  # this is necessary to later close excel  # removed 11-05-20 as feature doesnt work

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

# edit csv to add client_name
    df = pd.read_csv(desired_csv_name_full)  # convert csv to df
    df['client_name'] = client_name  # add col to df where title = 'client_name' and every cell says the client name
    df.to_csv(desired_csv_name_full, index=None)  # overwrite the csv file with the version with the new column

# clone and rename xlsx files from template
    fname_template = f'Summary {p_number} {se_general.get_shortened_str(survey_name, 10)}.xlsx'
    xls_final_filename_full = f"{project_dir_full}\\{fname_template}"
    print(f'attempting to copy {cfg.xls_template_full} to {xls_final_filename_full}\\')
    shutil.copy(cfg.xls_template_full, xls_final_filename_full)

    # add_client_name_to_xls()  # commented out as it corrupts the xlsx

# open all files

    p = subprocess.Popen(f'explorer "{xls_final_filename_full}"')
    time.sleep(20)
    se_general.excel_refresh_all()  # another refresh attempt
    # time.sleep(5)
    # if close_excel:
        # excel.Visible = True
        # print('attempting to quit excel')
        # excel.Application.Quit()
        # TODO: figure out how to get this to work

# pop up dialog box once script finished
gui.popup_finished_message()
