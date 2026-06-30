# Excel Online Office Scripts Setup

1. Upload `outputs/recruiter_notetaking_helper/Recruiter_Notetaking_Helper_v3.xlsx` to the shared SharePoint or OneDrive folder.
2. Keep the sheets `hd_employees`, `hd_requisitions`, and `input_catalog`; tables `tblHdEmployees`, `tblHdRequisitions`, and `tblInputCatalog`; and their column headers unchanged. `tblHdEmployees` includes both `recent_hire_date` and `date_of_last_mobility_event`; Intake displays the mobility-event date and stores recent hire date in the submitted notes table. The weekly Python jobs should replace rows in the `hd_` tables only. Do not replace `tblCandidateNotes`; it stores the conversation history.
3. Open the workbook in Excel for the web with a Microsoft 365 business or education account.
4. Go to `Automate`, create a new script for each file below, and paste the matching code:
   - `LookupEmployee.ts`
   - `OpenCandidateNotes.ts`
   - `LookupRequisition.ts`
   - `LookupEmployee_Pipeline.ts`
   - `AddDesiredRole.ts`
   - `AddDesiredLevel.ts`
   - `AddDesiredFunction.ts`
   - `AddSkill.ts`
   - `DeleteRole.ts`
   - `DeleteLevel.ts`
   - `DeleteFunction.ts`
   - `DeleteSkill.ts`
   - `SubmitCandidate.ts`
   - `SubmitCandidate_Pipeline.ts`
   - `ResetIntakeForm.ts`
5. Optional: add Automate buttons for those fifteen scripts. Suggested labels are `Lookup Employee`, `Open Candidate Notes`, `Lookup Requisition`, `Lookup Pipeline Candidate`, `Add Desired Role`, `Delete Role`, `Add Desired Level`, `Delete Level`, `Add Desired Function`, `Delete Function`, `Add Skill`, `Delete Skill`, `Submit Notes`, `Submit Pipeline Candidate`, and `Reset Form`. Place `Open Candidate Notes` near the conversation timeline. Place `Lookup Pipeline Candidate` and `Submit Pipeline Candidate` on `Pipeline`.
6. Set the default screener name in `Settings!C16`. This is used when the form resets because Office Scripts cannot reliably read the signed-in user's display name.
7. Test with `100245` or `Avery Martinez` to see multiple demo conversations in the timeline. Test requisition lookup with `REQ-2026-0142` or `FP&A Manager`.
8. The Screen date field is at the top of the form, blank, and date-formatted so Excel Online presents its date picker. If left blank, `SubmitCandidate.ts` saves today's date. Every submitted conversation also receives a precise `added_datetime` timestamp, while the timeline sorts by `screen_date`.
9. Requisition lookup also populates Job level. On `Intake`, Candidate applied and Candidate stage are recruiter-entered dropdowns; edit their allowed values in `Settings`. The same `LookupRequisition.ts` script works on `Intake` and `Pipeline`; on `Pipeline`, it stores Job level behind the scenes and refreshes the visible internal candidate list from `tblCandidateNotes`.
10. Employee lookup restores the most recent desired-role, desired-level, desired-function, and skills chains. `AddDesiredRole.ts`, `AddDesiredLevel.ts`, `AddDesiredFunction.ts`, and `AddSkill.ts` check the catalog and historical submitted values before adding. Exact matches and clear new entries append immediately. For a fuzzy match scoring 80 or higher, a hidden Similar row appears directly below that field; choose an option and press the same add button again. The row hides after the value is added. `DeleteRole.ts`, `DeleteLevel.ts`, `DeleteFunction.ts`, and `DeleteSkill.ts` remove an exact normalized input value from the matching chain; if the value is not present, the same hidden row appears with a not-found message.
11. On `Pipeline`, run `LookupRequisition.ts` first, then run `LookupEmployee_Pipeline.ts` for the employee to add. `SubmitCandidate_Pipeline.ts` writes a preliminary `tblCandidateNotes` row with candidate and requisition details, leaves `screen_date` blank, clears only the candidate-add fields, keeps the selected requisition displayed, and refreshes the pipeline list.
12. If a submit script is missing required information, it writes the missing field list directly to the Intake status strip or Pipeline status field before stopping.
13. Run `OpenCandidateNotes.ts` after employee lookup to activate the `Candidates` sheet and pre-filter `tblCandidateNotes` to the loaded candidate. It filters by employee_id first, then mm_id, then full name if needed.
14. `tblInputCatalog` starts with common skills, roles, levels, and functions but is designed to grow from user entries. The `type` column separates `skill`, `role`, `level`, and `function` rows.
15. The `Intake` sheet is protected so recruiters edit the intended input cells only. The scripts temporarily unprotect the sheet when they need to populate lookup outputs or reset the form.

The `Intake` sheet intentionally avoids embedded setup instructions so it stays clean for recruiter demos and day-to-day use.
