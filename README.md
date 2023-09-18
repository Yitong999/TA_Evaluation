# TA Evaluation tools

## Objective
Currently Chemistry department record TA evaluations using Qualtrics. 
They then take these evaluations and have to summarize them into
individual TA forms manually. This repository is to figure out
a program that would be able to take the data from Qualtrics 
(as a CSV file) and output the individual TA evaluations. There is
a big backlog with this and automation would make a world of difference.


## Setup python required library
```bash
pip install xlrd
pip install openpyxl
```

## How to use the tool
1. Make sure to save all useful files in **Output** folder before
   run the program again. The program will delete all previous files
   in **Output** folder.
1. Put all excel files in **Work_Space** folder.
2. Run the command below.
    ```bash
    python main.py
    ```
   Then, all results are stores in **Output** folder
