<H1>Incident and Changes Analyzer</H1>

The purpose of this set of scripts is to analyze and compare a list of key words with ticket descriptions columns in order to search for typos. The script compares strings using Levenshtein Distance and returns a percentage of coincidence.
Any coincidence of 80% or more will be impacted on an empty cell in the Excel file.

Usage: python3 <yourfile.py> -i inputfile -o outputfile

Note: You can use the same input and output file, it will work the same
