# Excel Online Office Scripts Setup

1. Upload `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v2.xlsx` to the shared SharePoint or OneDrive folder.
2. Keep the sheet `hd_employees`, table `tblHdEmployees`, and its column headers unchanged. The weekly Python roster job should replace rows in that table only.
3. Open the workbook in Excel for the web with a Microsoft 365 business or education account.
4. Go to `Automate`, create a new script for each file below, and paste the matching code:
   - `LookupEmployee.ts`
   - `SubmitCandidate.ts`
   - `ResetIntakeForm.ts`
5. Optional: add Automate buttons for those three scripts. Suggested labels are `Lookup Employee`, `Submit Notes`, and `Reset Form`.
6. Test with a demo lookup such as `100245`, `MM-008512`, or `Avery Martinez`.

The `Intake` sheet intentionally avoids embedded setup instructions so it stays clean for recruiter demos and day-to-day use.
