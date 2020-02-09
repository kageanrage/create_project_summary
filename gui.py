import PySimpleGUI as sg


def get_inputs_via_gui():
    layout = [
        [sg.Checkbox('All live projects', default=True, key="include_all_live_jobs")],
        [sg.Text('Surveys to exclude', size=(20, 1), justification='center', font=("Helvetica", 15))],
        [sg.Checkbox('Welcome survey', size=(10, 1), default=True, key="welcome_survey"),
         sg.Checkbox('Education survey', default=True, key="education_survey")],
        [sg.Checkbox('Member sat survey', default=True, key="member_survey")],

        [sg.Text('Manual inclusions', size=(30, 1), justification='center', font=("Helvetica", 15))],
        [sg.Text('Survey 1', size=(25, 1)), sg.InputText('', key="survey_1")],
        [sg.Text('Survey 2', size=(25, 1)),
         sg.InputText('', key="survey_2")],
        [sg.Text('Survey 3', size=(25, 1)), sg.InputText('', key="survey_3")],
        [sg.Text('Survey 4', size=(25, 1)), sg.InputText('', key="survey_4")],
        [sg.Submit(), sg.Cancel()]
    ]

    window = sg.Window('Create Project Summaries').Layout(layout)

    button, values = window.Read()

    include_all_live_jobs = values['include_all_live_jobs']
    print(f'include_all_live_jobs = {include_all_live_jobs}')

    if include_all_live_jobs:
        manually_select_projects = False
    else:
        manually_select_projects = True

    print(f"values['welcome_survey'] = {values['welcome_survey']}")
    print(f"values['education_survey'] = {values['education_survey']}")
    print(f"values['member_survey'] = {values['member_survey']}")


    # manual_survey_1 = values['survey_1']
    # manual_survey_2 = values['survey_2']
    # manual_survey_3 = values['survey_3']
    # manual_survey_4 = values['survey_4']

    # section: this will look at the exclusion checkboxes and create a list
    surveys_to_exclude = []
    potential_exclusions = ['welcome_survey', 'education_survey', 'member_survey']
    for potential_exclusion in potential_exclusions:
        if values[f'{potential_exclusion}']:
            surveys_to_exclude.append(f'{potential_exclusion}')

    print('surveys_to_exclude looks like this:')
    print(surveys_to_exclude)

    # section: this will look at the manual inclusions text entry boxes and create a list
    manual_inclusions = []
    for i in range(1, 5):
        if values[f'survey_{i}'] != "":
            manual_inclusions.append(values[f'survey_{i}'])

    print('manual_inclusions looks like this:')
    print(manual_inclusions)

    window.close()
    return manually_select_projects, surveys_to_exclude, manual_inclusions


def popup_finished_message():
    sg.Popup('Script finished')  # Shows OK button




def main():
    all_live_jobs, surveys_to_exclude, manual_inclusions = get_inputs_via_gui()


if __name__ == '__main__':
    main()
