import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import messagebox


def get_inputs_via_gui():
    """
    Display a modern GUI form to get user inputs for project summary creation.
    Returns: (manually_select_projects, surveys_to_exclude, manual_inclusions) or None if cancelled
    """
    # Create the main window with a modern theme
    root = ttk.Window(themename="cosmo")  # You can change to "flatly", "darkly", "superhero", etc.
    root.title("Create Project Summaries")
    root.geometry("550x520")
    root.resizable(True, True)
    root.minsize(500, 500)
    
    # Center the window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Variables to store values
    include_all_live_jobs = ttk.BooleanVar(value=True)
    welcome_survey = ttk.BooleanVar(value=True)
    iris_recruit = ttk.BooleanVar(value=True)
    survey_vars = {
        'survey_1': ttk.StringVar(),
        'survey_2': ttk.StringVar(),
        'survey_3': ttk.StringVar(),
        'survey_4': ttk.StringVar()
    }
    
    result = None  # Will store the return value
    
    def on_submit():
        nonlocal result
        
        include_all = include_all_live_jobs.get()
        
        if include_all:
            manually_select_projects = False
        else:
            manually_select_projects = True
        
        print(f'include_all_live_jobs = {include_all}')
        print(f"welcome_survey = {welcome_survey.get()}")
        print(f"iris_recruit = {iris_recruit.get()}")
        
        # Build surveys_to_exclude list
        surveys_to_exclude = []
        if welcome_survey.get():
            surveys_to_exclude.append('Welcome Survey')
        if iris_recruit.get():
            surveys_to_exclude.append('Iris Recruit Apr-22')
        
        print('surveys_to_exclude looks like this:')
        print(surveys_to_exclude)
        
        # Build manual_inclusions list
        manual_inclusions = []
        for i in range(1, 5):
            survey_value = survey_vars[f'survey_{i}'].get().strip()
            if survey_value:
                manual_inclusions.append(survey_value)
        
        print('manual_inclusions looks like this:')
        print(manual_inclusions)
        
        result = (manually_select_projects, surveys_to_exclude, manual_inclusions)
        root.destroy()
    
    def on_cancel():
        nonlocal result
        result = None
        root.destroy()
    
    # Create main frame with padding
    main_frame = ttk.Frame(root, padding=20)
    main_frame.pack(fill=BOTH, expand=True)
    
    # Title section
    title_label = ttk.Label(
        main_frame, 
        text="Project Summary Configuration", 
        font=("Helvetica", 16, "bold")
    )
    title_label.pack(pady=(0, 20))
    
    # All live projects checkbox
    all_live_check = ttk.Checkbutton(
        main_frame,
        text="All live projects",
        variable=include_all_live_jobs,
        bootstyle="primary-round-toggle"
    )
    all_live_check.pack(anchor=W, pady=5)
    
    # Separator
    ttk.Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=15)
    
    # Surveys to exclude section
    exclude_label = ttk.Label(
        main_frame,
        text="Surveys to exclude",
        font=("Helvetica", 12, "bold")
    )
    exclude_label.pack(anchor=W, pady=(0, 10))
    
    welcome_check = ttk.Checkbutton(
        main_frame,
        text="Welcome Survey",
        variable=welcome_survey,
        bootstyle="primary-round-toggle"
    )
    welcome_check.pack(anchor=W, pady=3)
    
    iris_check = ttk.Checkbutton(
        main_frame,
        text="Iris Recruit Apr-22",
        variable=iris_recruit,
        bootstyle="primary-round-toggle"
    )
    iris_check.pack(anchor=W, pady=3)
    
    # Separator
    ttk.Separator(main_frame, orient=HORIZONTAL).pack(fill=X, pady=15)
    
    # Manual inclusions section
    manual_label = ttk.Label(
        main_frame,
        text="Manual inclusions",
        font=("Helvetica", 12, "bold")
    )
    manual_label.pack(anchor=W, pady=(0, 10))
    
    # Survey input fields
    for i in range(1, 5):
        survey_frame = ttk.Frame(main_frame)
        survey_frame.pack(fill=X, pady=3)
        
        label = ttk.Label(survey_frame, text=f"Survey {i}:", width=12, anchor=W)
        label.pack(side=LEFT, padx=(0, 10))
        
        entry = ttk.Entry(survey_frame, textvariable=survey_vars[f'survey_{i}'], width=30)
        entry.pack(side=LEFT, fill=X, expand=True)
    
    # Button frame
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill=X, pady=(25, 10))
    
    submit_btn = ttk.Button(
        button_frame,
        text="Submit",
        command=on_submit,
        bootstyle="success",
        width=12
    )
    submit_btn.pack(side=RIGHT, padx=(10, 0))
    
    cancel_btn = ttk.Button(
        button_frame,
        text="Cancel",
        command=on_cancel,
        bootstyle="secondary",
        width=12
    )
    cancel_btn.pack(side=RIGHT)
    
    # Handle window close (X button)
    root.protocol("WM_DELETE_WINDOW", on_cancel)
    
    # Run the GUI
    root.mainloop()
    
    return result


def popup_finished_message():
    """
    Display a popup message indicating the script has finished.
    """
    root = ttk.Window(themename="cosmo")
    root.withdraw()  # Hide the main window
    
    messagebox.showinfo(
        title="Script Complete",
        message="Script finished",
        parent=root
    )
    
    root.destroy()


def main():
    all_live_jobs, surveys_to_exclude, manual_inclusions = get_inputs_via_gui()


if __name__ == '__main__':
    main()
