# Investigator Notebook

An Excel/VBA workbook that sets up a consistent case file for financial
crime investigations. The investigator pastes the account numbers under
review, clicks one button, and gets:

- **A tab for every account**, named with the account number, containing:
  - a structured **ACCOUNT INFO** block (name, signers, segment, account
    and product type, balance, open and closing dates), with dropdowns
    where values come from a fixed list
  - a **summary area** for documenting account activity
  - a **HOME** button to return to the case overview
- **A case overview on the HOME tab** that lists every account with:
  - its name, pulled automatically from the account tab
  - a link to its tab
  - an "Is suspicious" flag
  - case number, review period and SAR disposition fields

> All account numbers and data used to demonstrate the workbook are
> fictional.

## Why it exists

Investigations often cover many accounts, and an unstructured workbook
means every investigator documents them differently. This workbook gives
each account the same layout and the same required fields, and the
dropdowns keep values consistent. The HOME tab then summarizes the whole
case on one screen.

## Features

- **Duplicate protection:** an account pasted twice, or already
  processed, never gets a second tab.
- **Input validation:** account numbers must be exactly 10 digits. Anything
  else is skipped and listed in a report after processing.
- **Batch-friendly:** run it again after pasting more accounts and the new
  ones are added to the existing list. A tab that was deleted by mistake is
  rebuilt.
- **Live account names:** the HOME list updates as soon as a name is typed
  on an account tab.
- **Placeholder text** in each summary area clears when you click into it
  and returns if the area is left empty.
- **Removing accounts:** a form lists every account tab. Double-click the
  ones to remove, confirm, and the notebook deletes their tabs, their rows
  in the HOME list, and their numbers in the paste area.

## Setup

1. Open `Notebook.xlsm` and enable macros.
2. Press **Alt + F11** to open the VBA editor.
3. **File > Import File** and choose `src/modNotebook.bas`.
4. Double-click **ThisWorkbook** in the project tree and paste in the
   contents of `src/ThisWorkbook.txt`.
5. **Insert > UserForm**. In the Properties window (**F4**), set its
   **(Name)** to `frmRemoveAccounts`. Right-click the form, choose
   **View Code**, and paste in the contents of `src/frmRemoveAccounts.txt`.
   The form builds its own lists and buttons, so leave it blank.
6. Back in Excel, press **Alt + F8**, run **SetupHome** once, then save.

## Using it

1. Paste up to 25 account numbers into **E2:E26** on the HOME tab.
2. Click **PROCESS ACCOUNT NUMBERS**.
3. Use the **Go to tab** links to open each account, fill in the ACCOUNT
   INFO block, and write the activity summary.
4. To remove accounts you don't need, click **REMOVE ACCOUNTS** on the
   HOME tab.

Line breaks within the summary area: **Alt + Enter**.

## Project layout

```
Notebook.xlsm            The workbook
src/modNotebook.bas      Main VBA module (exported so the code is readable on GitHub)
src/ThisWorkbook.txt     Workbook event code for the summary placeholder
src/frmRemoveAccounts.txt  Code for the Remove Accounts form
```
