# markcount

Mark exam sections for each student and save results in a file.

This program is specifically tailored for IdC’s placement test.

## What does it do?

`markcount.py` parses information from `marks.xlsx` and `answerKey.xlsx`, counts students’ marks, compares them against the answer key, tallies them, and outputs the results in `results.xlsx`.

## How to use it

### Python

Make sure [Python](https://www.python.org) and `pip` are installed. This can easily be checked by running these two commands:

```
python --version
pip --version
```

Depending on your machine, e.g., macOS, it might be necessary to replace `python` by `python3` and `pip` by `pip3`.

### Required files

Make sure to place all relevant files in the same folder:

* `markcount.py`
* `answerKey.xlsx`
* `marks.xlsx`


### Libraries

This program relies on the `openpyxl` library to read and write Microsoft Excel files. 

#### The recommended way

The recommended way is to use `pipenv` to manage this program’s libraries.

#### The straightforward way

Install `openpyxl` globally by running `pip install openpyxl`. Again, it might be necessary to replace `pip` by `pip3` on some machines.

### Running the program

```
python markcount.py
```

## Possible updates

This program currently works for only one type of placement test. It is possible to update it so it could work for other types with a different number of sections.
