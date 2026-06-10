# Excel Online Recruiter Notetaking Helper

Polished Microsoft Excel workbook for internal recruiting screens in Microsoft 365 online.

The workbook uses Excel as the recruiter-facing front end. Recruiters can look up an internal employee from a weekly refreshed roster, capture screening notes, score rubric categories, and submit the completed screen to a structured candidate-notes table.

## Workbook Structure

- `Intake`: clean recruiter interface for employee lookup and screening notes.
- `Candidates`: structured `tblCandidateNotes` table storing submitted notes.
- `hd_employees`: weekly refreshed `tblHdEmployees` roster cache populated by an external Python job.
- `Settings`: dropdown values and score meanings.
- `Instructions`: concise workbook setup and usage notes.

## Office Scripts

- `office-scripts/LookupEmployee.ts`: searches `tblHdEmployees` by full name, `employee_id`, or `mm_id` and populates Candidate Details on `Intake`.
- `office-scripts/SubmitCandidate.ts`: validates required intake fields and appends the screen to `tblCandidateNotes`.
- `office-scripts/ResetIntakeForm.ts`: clears editable intake fields without changing labels, dropdowns, formatting, or data tables.

## Files

- `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v2.xlsx`
- `office-scripts/LookupEmployee.ts`
- `office-scripts/SubmitCandidate.ts`
- `office-scripts/ResetIntakeForm.ts`
- `office-scripts/SETUP.md`

## Excel Online Setup

1. Upload the workbook to OneDrive or SharePoint.
2. Open it in Excel for the web with a Microsoft 365 business or education account.
3. In the `Automate` tab, create scripts from the TypeScript files in `office-scripts`.
4. Optionally add Automate buttons for lookup, submit, and reset actions.
5. Keep the `hd_employees` sheet/table headers stable so the weekly Python process and lookup script continue to work.

## Contribution Policy

This repository is public for viewing and reference only. External edits, issues, pull requests, or contribution requests are not accepted.
