#sustainable
A project to automate the report created by the UN Sustainable Development goals for a Metropolitan-Statistical-Area (MSA)

It generates a template doucment with a summary figure and the detailed indicators in LaTeX format

Then you can open the created project in a LaTeX editor (like TeXstudio), edit it, and generate the PDF



# How-To
1. download the entire repository
2. determine which MSA you want to create the report for (look at the excel file, in the data dir)
3. edit the make_report.py program to pick that one (put in the name of the maincity in the "main" section
4. run the make_report.py program via python 3  `python3 make_report.py`
5. edit the generated report in a LaTeX editor, to add comments edit the MSA-sdg.tex file and the individual files found in the comments directory (1 file for each indicator).
