# Excel Online Recruiter Notetaking Helper

Polished Microsoft Excel workbook for recruiter interview screens in Microsoft 365 online.

The workbook includes:

- `Intake` form sheet for candidate notes and recruiter-screen rubric scores
- `Candidates` table for submitted candidate records
- `Dashboard` with pipeline, rubric, and next-step summaries
- `Settings` sheet for dropdown values and score definitions
- Office Scripts TypeScript files for Excel Online automation

## Files

- `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper.xlsx`
- `office-scripts/SubmitCandidate.ts`
- `office-scripts/ResetIntakeForm.ts`
- `office-scripts/RefreshDashboard.ts`
- `office-scripts/FormatWorkbook.ts`
- `office-scripts/SETUP.md`

## Excel Online Setup

1. Upload the workbook to OneDrive or SharePoint.
2. Open it in Excel for the web with a Microsoft 365 business or education account.
3. In the `Automate` tab, create scripts from the TypeScript files in `office-scripts`.
4. Run `FormatWorkbook.ts` once.
5. Add script buttons on the `Intake` sheet for submit, reset, and dashboard refresh actions.

## Contribution Policy

This repository is public for viewing and reference only. External edits, issues, pull requests, or contribution requests are not accepted.
