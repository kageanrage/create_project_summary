#######
# Purpose of this is to download a number of per-project surveydata CSV files so they can be merged into the Historical Survey Data csv
# This is needed because if the Data Model updates aren't run regularly, there can form gaps in the data
# Initially this will be a stand alone .py file but ultimately I should combine it with the script which merges the csv files together
#####


# I'll lay out the steps to take:

# Using create_project_summary.py as a reference, grab the list of all project from survey admin
# Take an input from the user which limits to projects with a Start Date of in the past X period of time
# For each project in that timeframe, download the CSV
# move the CSVs to a directory