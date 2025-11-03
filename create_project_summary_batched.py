import se_admin, se_general
from config import Config
import os, subprocess, datetime, shutil, time, sys
import pandas as pd
from pprint import pprint
import gui
import win32com.client as win32
import win32gui
import win32con


cfg = Config()


def match_csv_to_project(p_number, downloads_dir, date_str):
    """
    Match CSV files to projects based on p_number and today's date.
    
    CSV filename format: _SE_SurveyResults_P-{p_number}_{date}_{time}.csv
    Example: _SE_SurveyResults_P-50244_3_11_2025 8_53_14 am.csv
    
    Matching criteria:
    1. Filename contains P-{p_number} (extract number after P-)
    2. Filename contains today's date in format d_M_yyyy (e.g., 3_11_2025)
    
    Args:
        p_number: Project number to match
        downloads_dir: Directory to scan for CSV files
        date_str: Today's date string in format d_M_yyyy
    
    Returns:
        Matching CSV filename or None if not found
    """
    # p_number already includes "P-" prefix (e.g., "P-50244"), so use it directly
    p_pattern = p_number
    
    # Scan downloads directory for matching CSV files
    try:
        csv_files = [f for f in os.listdir(downloads_dir) if f.endswith('.csv')]
    except FileNotFoundError:
        print(f"Downloads directory not found: {downloads_dir}")
        return None
    
    # Find files matching both criteria
    for csv_file in csv_files:
        # Check if filename contains the p_number pattern (e.g., "P-50244")
        if p_pattern in csv_file:
            # Check if filename contains today's date
            if date_str in csv_file:
                return csv_file
    
    return None


# determine if running in auto mode (via BAT file) or not (using GUI for manual inputs)
if len(sys.argv) > 1:
    print('running via script in auto mode - ')
    manually_select_projects = False
    print(f'manually_select_projects = {manually_select_projects}')
    surveys_to_exclude = ['Welcome Survey', 'IrisRecruitApr22']
    print('surveys_to_exclude = ')
    print(surveys_to_exclude)
    manual_inclusions = []
    print('manual_inclusions:')
    print(manual_inclusions)
    # Record script start time for auto mode
    script_start_time = datetime.datetime.now()
else:
    gui_results = gui.get_inputs_via_gui()
    if gui_results is None: # Check if GUI returned None (cancelled)
        print("User cancelled the operation from the GUI. Exiting script.")
        sys.exit(0)
    manually_select_projects, surveys_to_exclude, manual_inclusions = gui_results
    # Record script start time when user clicks Submit/Go in GUI
    script_start_time = datetime.datetime.now()
    print(f"\nScript started at: {script_start_time.strftime('%Y-%m-%d %H:%M:%S')}")




# ==============================================================================
# NORMAL MODE: Initialize browser and get project data
# ==============================================================================
# scrape the latest survey admin table

driver = se_general.init_selenium()
driver.implicitly_wait(30)
live_jobs, all_jobs = se_admin.grab_sa_projects_dicts(driver)


# create dir for todays date
root_dir = cfg.project_updates_root_dir
todays_date = str(datetime.datetime.today())[0:10].replace('-', '')
today = str(datetime.datetime.today().strftime("%A %d. %B %Y"))[0:3]
date_dir_name = f"{todays_date} {today}"
se_general.create_dir_if_not_exists(f"{root_dir}\\{date_dir_name}")

# Get today's date in format d_M_yyyy for CSV matching (e.g., 3_11_2025)
today_obj = datetime.datetime.today()
date_str = f"{today_obj.day}_{today_obj.month}_{today_obj.year}"

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
    # Remove excluded surveys from live_jobs before running main loop
    potential_exclusions = ['Welcome Survey', 'Iris Recruit Apr-22']
    for potential_exclusion in potential_exclusions:
        if potential_exclusion in surveys_to_exclude and jobs_of_interest[f'{potential_exclusion}'] != False:
            del jobs_of_interest[f'{potential_exclusion}']
            print(f"Deleting {potential_exclusion} from jobs_of_interest")

print('Commencing loop for jobs_of_interest, which looks like this:')
pprint(jobs_of_interest)


# ==============================================================================
# PHASE 1: Download all CSVs
# ==============================================================================
print('\n========== PHASE 1: Downloading all CSVs ==========')

projects_metadata = []  # Store project metadata for later phases

for k in jobs_of_interest.keys():
    print(f"\nDownloading CSV for jobs_of_interest['{k}']")

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
    print(f'\nNow downloading SurveyResults for {dl_link}')
    driver.get(dl_link)
    
    # Minimal wait since downloads are instant
    time.sleep(0.5)

    # Store project metadata for later phases
    projects_metadata.append({
        'p_number': p_number,
        'client_name': client_name,
        'survey_name': survey_name,
        'survey_id': survey_id,
        'project_dir_full': project_dir_full,
        'project_dir_name': project_dir_name,
        'project_subdir_full': project_subdir_full
    })

# Wait a moment for all downloads to complete before matching
print('\nWaiting for all downloads to complete...')
time.sleep(2)

# Now match CSV files to projects after all downloads complete
print('Matching CSV files to projects...')
for project in projects_metadata:
    p_number = project['p_number']
    most_recent_csv = match_csv_to_project(p_number, cfg.downloads_dir, date_str)
    
    if most_recent_csv:
        project['csv_filename'] = most_recent_csv
        project['csv_full_path'] = f"{cfg.downloads_dir}\\{most_recent_csv}"
        print(f'Matched CSV {most_recent_csv} to project {p_number}')
    else:
        print(f'WARNING: Could not find CSV for project {p_number}')
        project['csv_filename'] = None
        project['csv_full_path'] = None

# Close browser/driver now that downloads are complete - no longer needed
if driver:
    print('\nClosing browser - downloads complete, no longer needed')
    driver.quit()
    driver = None

# Pre-emptively open Excel application so it's ready when we start opening files
# This prevents the first file from not refreshing due to Excel initialization time
print('\nPre-emptively opening Excel application...')
excel_app = None
try:
    excel_app = win32.Dispatch('Excel.Application')
    excel_app.Visible = True
    excel_app.DisplayAlerts = False  # Suppress alerts
    
    # Create a temporary workbook to ensure Excel has a visible window
    try:
        # Check if there are any workbooks, if not create one
        if excel_app.Workbooks.Count == 0:
            wb_temp = excel_app.Workbooks.Add()
        # Get the main Excel window
        excel_app.WindowState = -4137  # xlMaximized constant (-4137)
        
        # Use Windows API to bring Excel window to foreground
        # Find Excel window by class name
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                # Excel main window class name
                if 'XLMAIN' in class_name or 'OpusApp' in class_name:
                    # Bring window to foreground
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
        
        # Enumerate all windows to find Excel
        win32gui.EnumWindows(enum_handler, None)
        
        # Also try to activate via Excel COM object
        try:
            if excel_app.ActiveWindow:
                excel_app.ActiveWindow.WindowState = -4137  # xlMaximized
                excel_app.ActiveWindow.Activate()
        except:
            pass
        
        print('Excel application opened, maximized, and brought to foreground')
        print('Excel will remain open during Phase 2 processing...')
    except Exception as e2:
        print(f"Warning: Could not properly maximize/bring Excel to foreground: {e2}")
        print("Excel is opened but may need manual activation")
except Exception as e:
    print(f"Warning: Could not pre-open Excel via win32com: {e}")
    print("Will attempt to open Excel when files are opened in Phase 3")
    excel_app = None

# ==============================================================================
# PHASE 2: Process all CSVs
# ==============================================================================
print('\n========== PHASE 2: Processing all CSVs ==========')

for project in projects_metadata:
    if project['csv_filename'] is None:
        print(f"Skipping project {project['p_number']} - no CSV file found")
        continue
    
    print(f"\nProcessing CSV for project {project['p_number']}")
    
    p_number = project['p_number']
    client_name = project['client_name']
    survey_name = project['survey_name']
    project_subdir_full = project['project_subdir_full']
    most_recent_csv_full = project['csv_full_path']
    most_recent_csv = project['csv_filename']

    # move downloaded results to appropriate dir, then rename
    shutil.move(most_recent_csv_full, project_subdir_full)
    moved_csv_full = f"{project_subdir_full}\\{most_recent_csv}"
    desired_csv_name_full = f"{project_subdir_full}\\SurveyResults.csv"
    os.rename(moved_csv_full, desired_csv_name_full)

    # edit csv to add client_name
    df = pd.read_csv(desired_csv_name_full)  # convert csv to df
    df['client_name'] = client_name  # add col to df where title = 'client_name' and every cell says the client name
    print('desired_csv_name_full =')
    print(desired_csv_name_full)
    print('Waiting for 1 sec before writing CSV')
    time.sleep(1)
    print('Now attempting to run to_csv function on the above path')
    df.to_csv(desired_csv_name_full, index=None)  # overwrite the csv file with the version with the new column

    # clone and rename xlsx files from template
    fname_template = f'Summary {p_number} {se_general.get_shortened_str(survey_name, 10)}.xlsx'
    xls_final_filename_full = f"{project['project_dir_full']}\\{fname_template}"

    print(f'attempting to copy {cfg.xls_template_simple_full} to {xls_final_filename_full}\\')
    shutil.copy(cfg.xls_template_simple_full, xls_final_filename_full)

    # Store Excel file path for Phase 3
    project['xls_final_filename_full'] = xls_final_filename_full

# Close the temporary Book1 workbook that was created to ensure Excel had a visible window
# This prevents it from being the active workbook during Phase 3
if excel_app and excel_app.Workbooks.Count > 0:
    try:
        # Find and close Book1 if it exists
        for i in range(excel_app.Workbooks.Count, 0, -1):
            wb = excel_app.Workbooks(i)
            if wb.Name.startswith('Book') and '.xlsx' not in wb.Name and '.xls' not in wb.Name:
                wb.Close(SaveChanges=False)
                print('Closed temporary Book1 workbook before Phase 3')
                break
    except Exception as e:
        print(f"Warning: Could not close temporary workbook: {e}")

# ==============================================================================
# PHASE 3: Open all Excel files using Approach 4 (Combined Delays + Manual Refresh)
# ==============================================================================
print('\n========== PHASE 3: Opening all Excel files ==========')
print("Using Approach 4: Combined delays and manual refresh")

first_file = True  # Track if this is the first file being opened

for project in projects_metadata:
    if 'xls_final_filename_full' not in project:
        print(f"Skipping project {project['p_number']} - no Excel file created")
        continue
    
    xls_final_filename_full = project['xls_final_filename_full']
    
    # Open file in pre-opened Excel application if available, otherwise use os.startfile
    if excel_app:
        try:
            # Open workbook in the pre-opened Excel application
            wb = excel_app.Workbooks.Open(xls_final_filename_full)
            print(f"Opened Excel file in pre-opened Excel application: {xls_final_filename_full}")
            # Activate the newly opened workbook window (not the temporary Book1)
            try:
                if wb.Windows.Count > 0:
                    wb.Windows(1).Activate()
            except:
                pass  # If activation fails, continue anyway
            # Note: Excel was already brought to foreground when first opened, no need to bring it to foreground again
        except Exception as e:
            print(f"Warning: Could not open file in pre-opened Excel app: {e}")
            print("Falling back to os.startfile method")
            try:
                os.startfile(xls_final_filename_full)
                print(f"Opened Excel file: {xls_final_filename_full}")
            except Exception as e2:
                print(f"Failed to open with os.startfile: {e2}")
                # Fallback to original method
                p = subprocess.Popen(f'explorer "{xls_final_filename_full}"')
                print("Fell back to explorer method")
    else:
        # Use os.startfile if Excel wasn't pre-opened
        try:
            os.startfile(xls_final_filename_full)
            print(f"Opened Excel file: {xls_final_filename_full}")
        except Exception as e:
            print(f"Failed to open with os.startfile: {e}")
            # Fallback to original method
            p = subprocess.Popen(f'explorer "{xls_final_filename_full}"')
            print("Fell back to explorer method")
    
    # Wait for Excel to fully load (10 seconds for first file, 5 seconds for others)
    if first_file:
        wait_time = 10
        print(f"Waiting {wait_time} seconds after opening first file before refresh...")
        first_file = False
    else:
        wait_time = 5
        print(f"Waiting {wait_time} seconds after opening before refresh...")
    time.sleep(wait_time)
    
    # Call manual refresh
    try:
        se_general.excel_refresh_all()
        print(f"Refreshed Excel file: {xls_final_filename_full}")
    except Exception as e:
        print(f"Warning: Could not refresh Excel file: {e}")
    
    # Wait 2 seconds before opening next file
    time.sleep(2)

# Driver was already closed after Phase 1, so no need to close again here

# Calculate and display script execution time
script_end_time = datetime.datetime.now()
script_duration = script_end_time - script_start_time

# Format duration nicely
hours, remainder = divmod(script_duration.total_seconds(), 3600)
minutes, seconds = divmod(remainder, 60)

if hours > 0:
    duration_str = f"{int(hours)}h {int(minutes)}m {int(seconds)}s"
elif minutes > 0:
    duration_str = f"{int(minutes)}m {int(seconds)}s"
else:
    duration_str = f"{int(seconds)}s"

print(f"\n{'='*60}")
print(f"Script execution completed in {duration_str}")
print(f"Start time: {script_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"End time: {script_end_time.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"{'='*60}")

# pop up dialog box once script finished
gui.popup_finished_message()
