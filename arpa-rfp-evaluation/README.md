# Scripts to Create and Process ARPA RFP Evaluations
The scripts in this directory are used to create copies of the 2021 ARPA RFP evaluation rubric to be used by evaluators and then to summarize the results of the evaluation. In the evaluation there were 70 proposals in response to the RFP, each of which was to be evaluated by 5 people, for a total of 350 evaluations.

To prepare to run, simply run the command: 
````
pip install --upgrade -r requirements.txt
````

There are four separate scripts, although the last three update separate tabs within a single output file:
- create_evaluator_sheets.py - Create evaluation spreadsheets for all evaluators and a spreadsheet with the evaluator/proposal/evaluation mappings.
- detail_reports.py - Create the _All Data_ and _Evaluation Status_ tabs in the summary report spreadsheet.
- summary_reports.py - Create the _Summary_ and _Potential Issues_ tabs in the summary report spreadsheet.
- check_stats.py - Create the _StatCheck_ tab in the summary report spreadsheet.

All input and output files are Google Sheets. IDs and other parameters are defined in the _inputs.json_ file.

## create_evaluator_sheets.py
This script uses the [master input spreadsheet](https://docs.google.com/spreadsheets/d/1uS7lPCDi28WzmIuPnnmbXvKb9bswv-wRcIjSF1NMY5k/edit#gid=0) to get evaluator assignments and the _README_ and evaluation rubric pages to use for each evaluator spreadsheet. It then creates one spreadsheet per evaluator, with a _README_ tab and one tab per assigned evaluation (each evaluation tab is a copy of the evaluation rubric page in the input spreadsheet, with information about the specific proposal inserted at the top). At the end of the process an _Evaluator Mappings_ is generated (see example [here](https://docs.google.com/spreadsheets/d/1KwLHE-qeyhEZDouAJqJumieCgiEVLXenAjFT6ranBOk/edit#gid=0)).

## detail_reports.py

This script uses the _Evaluator Mappings_ file generated above to read all evaluations and output a tab with the status of each evaluation (_Evaluation Status_ tab) as either _Complete_, _Partial_ or _Not Started_, and a tab with data on each individual question of each evaluation (_All Data_ tab). See sample [here](https://docs.google.com/spreadsheets/d/1PdRBJUuzohHSKxq1dg_WA_g9xFoQ_f2StrRZUHG36rw/edit#gid=0).

## summary_reports.py

This script uses the _Evaluator Mappings_ file generated above as well as the _Score Weighting_ information from the [master input spreadsheet](https://docs.google.com/spreadsheets/d/1uS7lPCDi28WzmIuPnnmbXvKb9bswv-wRcIjSF1NMY5k/edit#gid=0) to compute overall scores of all the proposals in the _Summary_ tab as well as a list of proposals with potential issues (e.g., a very wide spread of individual scores) in the _Potential Issues_ tab. Only fully completed evaluations are included in calculating these statistics.

## check_stats.py

This script computes score statistics for all evaluations using a different approach than above and then compares the two results as a check for correctness.

