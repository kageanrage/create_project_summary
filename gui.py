import PySimpleGUI as sg


def get_inputs_via_gui():
    layout = [
        [sg.Radio('Simple Report (no RR data)', "simple_or_complex", key='simple',  default=True, size=(10, 1)),
         sg.Radio('Complex Report (RR data)', "simple_or_complex", key='complex')],
        [sg.Checkbox('All live projects', default=False, key="include_all_live_jobs")],
        [sg.Text('Surveys to exclude', size=(20, 1), justification='center', font=("Helvetica", 15))],
        [sg.Checkbox('Welcome survey', size=(10, 1), default=True, key="Welcome Survey"),
         sg.Checkbox('Iris Recruit Apr-22', default=True, key="Iris Recruit Apr-22")],

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

    # Handle cancellation - Check for None, sg.Cancel object, sg.WIN_CLOSED, or the string 'Cancel'
    if button == sg.Cancel or button is None or button == sg.WIN_CLOSED or button == 'Cancel':
        window.close()
        return None # Indicate cancellation to the caller

    include_all_live_jobs = values['include_all_live_jobs']

    print(f'include_all_live_jobs = {include_all_live_jobs}')

    if include_all_live_jobs:
        manually_select_projects = False
    else:
        manually_select_projects = True

    if values['complex']:
        complex_report = True
    else:
        complex_report = False

    print(f"values['Welcome Survey'] = {values['Welcome Survey']}")
    print(f"values['Iris Recruit Apr-22'] = {values['Iris Recruit Apr-22']}")
    print(f"values['simple'] = {values['simple']}")
    print(f"values['complex'] = {values['complex']}")


    # manual_survey_1 = values['survey_1']
    # manual_survey_2 = values['survey_2']
    # manual_survey_3 = values['survey_3']
    # manual_survey_4 = values['survey_4']

    # section: this will look at the exclusion checkboxes and create a list
    surveys_to_exclude = []
    potential_exclusions = ['Welcome Survey', 'Iris Recruit Apr-22']
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
    return manually_select_projects, surveys_to_exclude, manual_inclusions, complex_report


def popup_finished_message():
    sg.Popup('Script finished')  # Shows OK button




def main():
    all_live_jobs, surveys_to_exclude, manual_inclusions = get_inputs_via_gui()


if __name__ == '__main__':
    main()
