# markcount

Mark exam sections for each student and saves results in a file.

This program is specifically tailored for IdC’s placement test.

## What does it do?

This program consists of two main files:

* `getanswerkey.py`: use it to generate an answer key file to use with the second program. Answer key should be provided as `answerKey.xlsx`. It creates `answerkey.py`, containing a Python variable.
* `markcount.py`: the main program that counts students’ marks and tallies them. Parses information from `marks.xlsx` and `answerkey.py` and outputs the results in `results.py`.

## How to use it

### Python

Make sure [Python](https://www.python.org) and `pip` are installed. This can easily be checked by running these two commands:

```
python --version
pip --version
```

Depending on your machine, e.g., macOS, it might be necessary to replace `python` by `python3` and `pip` by `pip3`.

This program relies on Python 3, not Python 2.

### Required files

Make sure to place all relevant files in the same folder:

* `markcount.py`
* `getanswerkey.py`
* `answerKey.xlsx`
* `marks.xlsx`


### Libraries

This program relies on the `openpyxl` library to read and write Microsoft Excel files. 

#### The recommended way

The recommended way is to use `pipenv` to manage this program’s libraries.

#### The straightforward way

Install `openpyxl` globally by running `pip install openpyxl`. Again, it might be necessary to replace `pip` by `pip3` on some machines.

### Running the program

1. Generate the answer key:

```
python getanswerkey.py
```

2. Generate the results file:

```
python markcount.py
```
