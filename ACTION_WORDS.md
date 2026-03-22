# Office Agent Action Words

This file explains how to use the program, how to give commands, and shows example commands for Excel, Word, and PowerPoint.

## How to use this program

1. Start the program with:
   `python main.py`
2. Leave it running in the background.
3. In any app or on any screen in Windows, type or paste a command.
4. Press `Enter` to execute it.

Command format:
`agent: <app>: <instruction>`

Supported apps:
- `excel`
- `word`
- `powerpoint`
- `ppt`

## How to give commands

You write commands in plain English, but they must start with `agent:` and include the target app.

Basic pattern:
- `agent: excel: <what you want Excel to do>`
- `agent: word: <what you want Word to do>`
- `agent: powerpoint: <what you want PowerPoint to do>`

Simple examples:
- `agent: excel: create a new workbook`
- `agent: excel: write 2500 in cell B3`
- `agent: word: create a new document`
- `agent: word: insert text Weekly sales summary`
- `agent: powerpoint: create a new presentation`
- `agent: powerpoint: add a new slide titled Project Update`

## Typing vs pasting

The program works in two ways:

1. Type a command anywhere, then press `Enter`
2. Copy a command, paste it anywhere, then press `Enter`

Important:
- Copying by itself does not execute anything.
- Pasting by itself does not execute anything.
- The command runs only after you press `Enter`.

Example:
- Copy: `agent: word: create a new document`
- Paste it into any text field
- Press `Enter`
- The program will execute the command

## Cross-app examples

You can control one Office app while working inside another app.

Examples:
- While in Excel:
  `agent: word: create a new document`
- While in Excel:
  `agent: word: open document C:/Users/faiza/Desktop/report.docx`
- While in Word:
  `agent: excel: create a new workbook`
- While in Word:
  `agent: excel: write 5000 in cell C4`
- While in PowerPoint:
  `agent: excel: open workbook C:/Users/faiza/Desktop/sales.xlsx`
- While in PowerPoint:
  `agent: word: insert text Meeting notes for March`

## Excel

Action words:
- `create_workbook`
- `open_workbook`
- `save_workbook`
- `save_workbook_as`
- `add_sheet`
- `set_active_sheet`
- `rename_sheet`
- `create_table`
- `write_cell`
- `write_formula`
- `write_range`
- `clear_range`
- `autofit_columns`

Examples:
- `agent: excel: create a new workbook`
- `agent: excel: open workbook C:/Users/faiza/Desktop/sales.xlsx`
- `agent: excel: save workbook`
- `agent: excel: save workbook as march_report.xlsx`
- `agent: excel: add a sheet named Summary`
- `agent: excel: switch to sheet Summary`
- `agent: excel: rename sheet Sheet1 to Revenue`
- `agent: excel: create a table with 5 rows and 4 columns starting at B2`
- `agent: excel: write 2500 in cell C7`
- `agent: excel: write formula =SUM(B2:B20) in C21`
- `agent: excel: clear range A1:D10`
- `agent: excel: autofit columns A to D`

## Word

Action words:
- `create_document`
- `open_document`
- `save_document`
- `save_document_as`
- `insert_text`
- `insert_table`
- `bold_headings`
- `find_replace`
- `set_alignment`

Examples:
- `agent: word: create a new document`
- `agent: word: open document C:/Users/faiza/Desktop/notes.docx`
- `agent: word: save document`
- `agent: word: save document as proposal.docx`
- `agent: word: insert text Quarterly summary goes here`
- `agent: word: insert a table with 3 rows and 5 columns`
- `agent: word: bold all headings`
- `agent: word: replace draft with final`
- `agent: word: align all text center`

## PowerPoint

Action words:
- `create_presentation`
- `open_presentation`
- `save_presentation`
- `save_presentation_as`
- `add_slide`
- `set_title`
- `insert_text_box`
- `insert_table`

Examples:
- `agent: powerpoint: create a new presentation`
- `agent: powerpoint: open presentation C:/Users/faiza/Desktop/demo.pptx`
- `agent: powerpoint: save presentation`
- `agent: powerpoint: save presentation as client_pitch.pptx`
- `agent: powerpoint: add a new slide titled Q2 Results`
- `agent: powerpoint: set title on slide 2 to Financial Overview`
- `agent: powerpoint: insert a text box on slide 1 with text Welcome Team`
- `agent: powerpoint: insert a table on slide 3 with 4 rows and 3 columns`

## Full command examples

Examples you can copy and use directly:

- `agent: excel: create a new workbook`
- `agent: excel: add a sheet named Budget`
- `agent: excel: write 12000 in cell B2`
- `agent: excel: write formula =SUM(B2:B10) in B11`
- `agent: excel: save workbook as budget_2026.xlsx`

- `agent: word: create a new document`
- `agent: word: insert text This is the project summary for March`
- `agent: word: insert a table with 4 rows and 3 columns`
- `agent: word: bold all headings`
- `agent: word: save document as summary_2026.docx`

- `agent: powerpoint: create a new presentation`
- `agent: powerpoint: add a new slide titled Sales Review`
- `agent: powerpoint: insert a text box on slide 1 with text Revenue increased this quarter`
- `agent: powerpoint: insert a table on slide 1 with 3 rows and 4 columns`
- `agent: powerpoint: save presentation as sales_review_2026.pptx`

## New text formatting actions

You can now control font size, color, and font family.

Excel examples:
- `agent: excel: set font size to 14 in range C3:H8`
- `agent: excel: change font color to blue in range C3:H8`
- `agent: excel: change font to Calibri in range C3:H8`

Word examples:
- `agent: word: create heading "Test completed" with font size 46`
- `agent: word: set font size to 20`
- `agent: word: set font color to blue`
- `agent: word: set font to Arial`

PowerPoint title vs subcontent:
- `title` means the main slide heading placeholder.
- `subtitle` means subtitle placeholder on title slides.
- `body` or `content` means the main text/content placeholder.

PowerPoint examples:
- `agent: powerpoint: set title on slide 1 to Project Update`
- `agent: powerpoint: set subtitle on slide 1 to Q1 Summary`
- `agent: powerpoint: set body text on slide 2 to Revenue increased by 15 percent`
- `agent: powerpoint: set title font size to 36 on slide 1`
- `agent: powerpoint: set subtitle font color to blue on slide 1`
- `agent: powerpoint: set body font to Calibri on slide 2`
