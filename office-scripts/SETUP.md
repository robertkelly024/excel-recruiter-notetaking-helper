# Excel Online Office Scripts Setup

1. Upload `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper.xlsx` to OneDrive or SharePoint.
2. Open it in Excel for the web with a Microsoft 365 business or education account.
3. Go to `Automate`, create a new script for each `.ts` file in this folder, and paste the matching code.
4. Run `FormatWorkbook.ts` once.
5. On the `Intake` sheet, add buttons for `SubmitCandidate.ts`, `ResetIntakeForm.ts`, and `RefreshDashboard.ts`.
6. Test the prefilled demo intake. `SubmitCandidate.ts` should append a row to `Candidates`, recalculate the dashboard, and clear the form.

Microsoft setup reference: https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons
