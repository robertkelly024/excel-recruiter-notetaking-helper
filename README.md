# Excel Online Recruiter Notetaking Helper

Polished Microsoft Excel workbook for internal recruiting screens in Microsoft 365 online.

The workbook uses Excel as the recruiter-facing front end. Recruiters can look up an internal employee and job requisition from weekly refreshed data tabs, review the employee's recent conversation timeline, maintain desired role, desired level, desired function, and skill chains, and submit each conversation to a structured candidate-notes table.

## Workbook Structure

- `Intake`: clean recruiter interface for employee lookup, requisition lookup, conversation history, and exploratory conversation notes.
- `Pipeline`: requisition-centered internal pipeline view for adding preliminary candidates and reviewing candidate status by requisition.
- `Candidates`: structured `tblCandidateNotes` table storing one timestamped row per conversation, including the referenced `hd_employees` and `hd_requisitions` source fields.
- `hd_employees`: weekly refreshed `tblHdEmployees` roster cache populated by an external Python job using the HR data field names required by the lookup scripts, including `preferred_full_nm`, `business_title_txt`, `recent_hire_dt`, `date_of_last_mobility_event`, manager fields, job profile fields, `year_nr`, and `month_nr`.
- `hd_requisitions`: weekly refreshed `tblHdRequisitions` requisition cache populated by an external Python job.
- `input_catalog`: reusable `tblInputCatalog` table for canonical skills, roles, desired levels, desired functions, aliases, usage counts, last-used dates, and status.
- `Settings`: dropdown values for next step, mobility interest, yes/no fields, candidate applied, and candidate stage.
- `Instructions`: concise workbook setup and usage notes.

## Office Scripts

- `office-scripts/LookupEmployee.ts`: searches `tblHdEmployees`, populates Candidate Details, restores the latest role/level/function/skill chains, and displays the five most recent conversations.
- `office-scripts/OpenCandidateNotes.ts`: activates `Candidates` and filters `tblCandidateNotes` to the currently loaded candidate.
- `office-scripts/LookupRequisition.ts`: searches `tblHdRequisitions` by `Requisition_ID` or `Job_Posting_Title` and populates Requisition Details, including job level, on `Intake` or `Pipeline`; on `Pipeline`, it also refreshes the internal candidate list.
- `office-scripts/LookupEmployee_Pipeline.ts`: searches `tblHdEmployees` and populates the Pipeline candidate-add fields.
- `office-scripts/AddDesiredRole.ts`: normalizes role input, checks `tblInputCatalog` plus historical submitted roles, and uses the compact fuzzy-match dropdown when needed.
- `office-scripts/AddDesiredLevel.ts`: applies the same catalog-backed add flow to desired levels.
- `office-scripts/AddDesiredFunction.ts`: applies the same catalog-backed add flow to desired functions.
- `office-scripts/AddSkill.ts`: normalizes skill input, checks `tblInputCatalog` plus historical submitted skills, appends exact matches and clear new skills, or temporarily reveals a compact field-specific dropdown for fuzzy matches scoring 80 or higher. The same script completes the selected option and hides the row on its next run.
- `office-scripts/DeleteRole.ts`, `DeleteLevel.ts`, `DeleteFunction.ts`, `DeleteSkill.ts`: remove an exact normalized input value from the corresponding pipe-delimited chain, or reveal that field's helper row when the value does not exist.
- `office-scripts/SubmitCandidate.ts`: validates required intake fields, shows any missing fields on the Intake status strip, and appends a row with `added_datetime` to `tblCandidateNotes`.
- `office-scripts/SubmitCandidate_Pipeline.ts`: validates required Pipeline fields, shows any missing fields in Pipeline status, adds a preliminary candidate row for the selected requisition, leaves `screen_date` blank, clears candidate fields, and refreshes the Pipeline list.
- `office-scripts/ResetIntakeForm.ts`: clears editable intake fields without changing labels, dropdowns, formatting, or data tables.
- `office-scripts/ResetPipelineForm.ts`: clears Pipeline lookup, quick-add, hidden source-cache, and list fields, then restores the default screener and ready status.

## Files

- `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v3.xlsx`
- `office-scripts/LookupEmployee.ts`
- `office-scripts/OpenCandidateNotes.ts`
- `office-scripts/LookupRequisition.ts`
- `office-scripts/LookupEmployee_Pipeline.ts`
- `office-scripts/AddDesiredRole.ts`
- `office-scripts/AddDesiredLevel.ts`
- `office-scripts/AddDesiredFunction.ts`
- `office-scripts/AddSkill.ts`
- `office-scripts/DeleteRole.ts`
- `office-scripts/DeleteLevel.ts`
- `office-scripts/DeleteFunction.ts`
- `office-scripts/DeleteSkill.ts`
- `office-scripts/SubmitCandidate.ts`
- `office-scripts/SubmitCandidate_Pipeline.ts`
- `office-scripts/ResetIntakeForm.ts`
- `office-scripts/ResetPipelineForm.ts`
- `office-scripts/SETUP.md`

## Excel Online Setup

1. Upload the workbook to OneDrive or SharePoint.
2. Open it in Excel for the web with a Microsoft 365 business or education account.
3. In the `Automate` tab, create scripts from the TypeScript files in `office-scripts`.
4. Optionally add Automate buttons for employee lookup, open candidate notes, requisition lookup, pipeline employee lookup, add/delete desired role, add/delete desired level, add/delete desired function, add/delete skill, submit intake notes, submit pipeline candidate, and reset actions.
5. Keep the `hd_employees` and `hd_requisitions` sheet/table headers stable so the weekly Python processes and lookup scripts continue to work.
6. Set the reusable screener default in `Settings!C16`. Excel's Office Scripts API does not reliably expose the signed-in user's display name.

## Contribution Policy

This repository is public for viewing and reference only. External edits, issues, pull requests, or contribution requests are not accepted.
