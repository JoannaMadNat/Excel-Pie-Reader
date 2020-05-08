# Excel Pie Reader
A simple program in Perl that reads data from a Windows Excel file and converts it to a PDF with a pie chart

## Usage
Run on linux using
```
$ perl main.pl file.xlsx [pdf_name]
```
 where 
 - file.xlsx is the Windows Excel file to be read
- optional: the name of the output PDF file. Default: output.pdf


## Example
```
$ perl main.pl thoughts.xlsx
```

## Features

### Pie Charts
To create a pie chart, write a consecutive series of rows containing the labels along with numbers in percentage format in the adjacent right-hand column. Program stops reading when it encounters an empty row.
ex:
| C1   |      C2      |  C3 |
|----------|:-------------:|------:|
| food |  30% |  |
| electricity |    20%   |    |
| gas | 50% |     |
|    | |     |



### Lists
To create a list, write a consecutive series of rows starting with the '-' character. program stops reading list when it encounters an empty row.
| C1   |      C2      |  C3 |
|----------|:-------------:|------:|
| - food |  |  |
| - electricity |       |    |
| - gas |  |     |
|  | |     |

### Title
The first populated row is always interpretted as the title and is formatted in bigger font on top of a line.
Coumn placement does not matter.

### Other Text
Any other text is placed in regular formatting where each cell occupies one line. The cells are read in row major order.
To place an empty line, populate a cell with the space ' ' character.