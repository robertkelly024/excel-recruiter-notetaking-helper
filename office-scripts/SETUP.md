# Excel Online Office Scripts Setup

1. Upload `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v3.xlsx` to the shared SharePoint or OneDrive folder.
2. Keep the sheets `hd_employees` and `hd_requisitions`, tables `tblHdEmployees` and `tblHdRequisitions`, and their column headers unchanged. The weekly Python jobs should replace rows in those tables only. Do not replace `tblCandidateNotes`; it stores the conversation history.
3. Open the workbook in Excel for the web with a Microsoft 365 business or education account.
4. Go to `Automate`, create a new script for each file below, and paste the matching code:
   - `LookupEmployee.ts`
   - `LookupRequisition.ts`
   - `ExpandCollapseJobDescription.ts`
   - `AddSkill.ts`
   - `SubmitCandidate.ts`
   - `ResetIntakeForm.ts`
5. Optional: add Automate buttons for those six scripts. Suggested labels are `Lookup Employee`, `Lookup Requisition`, `Expand/Collapse Job Description`, `Add Skill`, `Submit Notes`, and `Reset Form`.
6. Set the default screener name in `Settings!C16`. This is used when the form resets because Office Scripts cannot reliably read the signed-in user's display name.
7. Test with `100245` or `Avery Martinez` to see multiple demo conversations in the timeline. Test requisition lookup with `REQ-2026-0142` or `FP&A Manager`.
8. The Screen date field defaults to today, accepts Excel's date picker, and is saved as `yyyy-mm-dd`. Every submitted conversation also receives a precise `conversation_datetime` timestamp.

The `Intake` sheet intentionally avoids embedded setup instructions so it stays clean for recruiter demos and day-to-day use.
