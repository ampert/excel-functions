# Excel Functions

This repository contains nice to have formulas and functions that can greatly speed up excel development.

## List of TODOs

1. toolModeOn() - lock, removes excel clutters like headers and ribbon
1. toolModeOff() - reverts toolmodeOn. best placed when workbook is closed. always save the workbook
1. filepathToCell(cell as string)
1. fileopenFromCell(cell as string)
1. copyFileFromCells(srcCell as string, destCell as string)
1. rowCopy(fromSheetName as string, fromStartCell as string, toSheetName as string, toStartCell as string)
1. columnCopy
1. columnToString(column)
1. columnLetterToInt(column)
1. CreateOutlookEmail(To, Cc, Bcc, Subject, Body as string, Attachment[] as string)
1. Two Way Lookup - Formula is =INDEX(value_lookup,MATCH(row_lookup_value,row_lookup_range,0),MATCH(col_lookup_value,col_lookup_range,0))

## Table of Contents

### Formulas

### Functions


### Old school Git

```git
    git add .
    git status
    git commit -m "Your commit message"
    git push
```