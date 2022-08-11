# ods-generate-table

## Description

A python script that generates an OpenDocument Spreadsheet Document format (ods extension) table aggregating data from input forms used to facilitate an initial screen of new prospective real estate opportunities in Japan. The table can be used, for example, to present a summary of monthly activity to relevant stakeholders.

## Requirements

* pyexcel 0.7.0               
* pyexcel-ezodf 0.3.4               
* pyexcel-io 0.6.6               
* pyexcel-ods 0.6.0               
* pyexcel-ods3 0.5.3               
* pyexcel-xlsx 0.6.0

## Installation
Step 1: clone repo

```
git clone https://github.com/az107hq/ods-generate-table.git
```
Step 2: Add env variables
```
export PATH_TO_SOURCE_FILES=path/to/source/files
export PATH_TO_TABLE_TEMPLATE=path/to/table/template
```

## Usage
```
python3 main.py
```

## Test
```
python3 test.py
```

## To Do
* [x] add check to filter out the deals from non-current periods
* [ ] add input fields for evaluation stages to screening template
* [ ] add income yield column to summary table

## References
* [pyexcel](http://www.pyexcel.org/)
* [python](https://www.python.org/)
