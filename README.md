# Excel Online Recruiter Notetaking Helper

Polished Microsoft Excel workbook for internal recruiting screens in Microsoft 365 online.

The workbook uses Excel as the recruiter-facing front end. Recruiters can look up an internal employee and job requisition from weekly refreshed data tabs, review the employee's recent conversation timeline, capture exploratory career conversation notes, collect skills, and submit each conversation to a structured candidate-notes table.

## Workbook Structure

- `Intake`: clean recruiter interface for employee lookup, requisition lookup, conversation history, and screening notes.
- `Candidates`: structured `tblCandidateNotes` table storing one timestamped row per conversation.
- `hd_employees`: weekly refreshed `tblHdEmployees` roster cache populated by an external Python job.
- `hd_requisitions`: weekly refreshed `tblHdRequisitions` requisition cache populated by an external Python job.
- `Settings`: dropdown values for stage, next step, mobility interest, and yes/no fields.
- `Instructions`: concise workbook setup and usage notes.

## Office Scripts

- `office-scripts/LookupEmployee.ts`: searches `tblHdEmployees` by full name, `employee_id`, or `mm_id`, populates Candidate Details, and displays the five most recent conversations from `tblCandidateNotes`.
- `office-scripts/LookupRequisition.ts`: searches `tblHdRequisitions` by `Requisition_ID` or `Job_Posting_Title` and populates Requisition Details on `Intake`.
- `office-scripts/ExpandCollapseJobDescription.ts`: toggles the job-description row between compact and expanded height.
- `office-scripts/AddSkill.ts`: appends one skill at a time into the pipe-delimited skills chain on `Intake`.
- `office-scripts/SubmitCandidate.ts`: validates required intake fields and appends a timestamped conversation to `tblCandidateNotes`.
- `office-scripts/ResetIntakeForm.ts`: clears editable intake fields without changing labels, dropdowns, formatting, or data tables.

## Files

- `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v3.xlsx`
- `office-scripts/LookupEmployee.ts`
- `office-scripts/LookupRequisition.ts`
- `office-scripts/ExpandCollapseJobDescription.ts`
- `office-scripts/AddSkill.ts`
- `office-scripts/SubmitCandidate.ts`
- `office-scripts/ResetIntakeForm.ts`
- `office-scripts/SETUP.md`

## Excel Online Setup

1. Upload the workbook to OneDrive or SharePoint.
2. Open it in Excel for the web with a Microsoft 365 business or education account.
3. In the `Automate` tab, create scripts from the TypeScript files in `office-scripts`.
4. Optionally add Automate buttons for employee lookup, requisition lookup, job-description expand/collapse, add skill, submit, and reset actions.
5. Keep the `hd_employees` and `hd_requisitions` sheet/table headers stable so the weekly Python processes and lookup scripts continue to work.
6. Set the reusable screener default in `Settings!C16`. Excel's Office Scripts API does not reliably expose the signed-in user's display name.

## Contribution Policy

This repository is public for viewing and reference only. External edits, issues, pull requests, or contribution requests are not accepted.
