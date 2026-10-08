Kohinoor workforce reporting app: upload a turnover workbook on `/`, review strength results, and download a formatted Excel report. `/config` manages the strength template.

## Report workflow

1. Upload one `.xlsx` or `.xls` workbook (up to 20 MB). Open or repair unreadable files in Excel and save a new, unencrypted copy.
2. Scanning detects turnover headers (Grade is optional) by name, in any order, after title rows. Complete tables on multiple worksheets and repeated headers are supported. Common header aliases, punctuation, case, whitespace, and explicit merged department/unit cells are handled. Incomplete sections and unreadable worksheets are flagged; irrecoverable workbook corruption produces an error.
3. Use Strength with its category sidebar and grouped subcategory tables (subcategory to designation), or Summary for Admin & HR subcategory totals and the five Spinning category totals. Summary rows drill into the relevant detailed table. Review validation notes. On-roll means the number of unique listed employees that match a template designation; the upload must represent the workforce you want to count. The provided columns contain no termination/status field, so the app does not infer departures or turnover rates.
4. Download a genuine `.xlsx` workbook with Summary and one sheet per configured category. Headers, subcategory sections and totals, two-row spacing below the heading, bounded column widths, formula-based vacancies, and the official logo are included. Positive vacancies are green, negative vacancies are red, and zero stays neutral in both Strength tables and Excel; workbook colours update through conditional formatting.

Matching resolves department/unit/subcategory first, then designation. Grade is display-only and is not checked for matching or duplicate conflicts; a missing Grade column or blank grade does not exclude employees. Known consolidated department names, common designation abbreviations and word order, punctuation and case are normalized. Numbered lines remain distinct; unresolved departments or designations are excluded with field-specific explanations. Template entries for the same category, subcategory and designation are combined in reports and their approved strength is summed without editing the JSON. Department + Unit can match a combined subcategory such as `SPINNING - RING - 1`. Category/Org Group can disambiguate candidates. Ambiguous and unmatched records are excluded from counts and listed for review. Employee number + Unit identifies duplicates; repeated records with the same identity and strength fields count once, even when grade, shift, rest day, payroll or joining date differs. Conflicting identities, departments or designations are excluded. Validation identifies counted duplicates separately from excluded records.

`../strength.json` at the workspace root is the single shared template for the whole app. Its path is defined once in `lib/strength-file.ts`; there is no separate seed copy. Run app and test commands from `strength-analysis`. Scanning and exports only read the template; only the configuration save action writes it. Designation, grade, cadre, and approved strength come from this file. Configuration status remains configuration-only. The template contains no individual employee names, IDs, gender, shifts, or joining dates: those are available in the employee details dialog.

Reports, employee records and case edits are stored only in this tab's sessionStorage. Reloading or navigating within the same tab restores the report; closing the tab clears it. New uploads always start with fresh cases. No previous-report JSON, cloud database, history comparison, background polling, or report-save server request is used. Scanning and exporting run in Web Workers. Session storage quota or access failures show a readable error and preserve the current report. Configuration remains in `../strength.json`; saving configuration locally still writes that file. On Vercel, template changes require updating the source JSON and redeploying because deployment files cannot be rewritten.

Branding uses the supplied updated `public/kohinoor-logo.png` directly in the navigation and every Excel worksheet. Downloads fetch this asset without browser caching and verify the latest root template before export.

Review Cases works at designation level. A raw vacancy of -3 creates three numbered cases under that designation. Each dropdown offers other designations in the same subcategory (Against Post), Social Security, or Other (type text). Other requires a reason. Every saved case offsets one negative vacancy; three saved cases bring -3 to 0. Approved strength and uploaded on-roll counts stay intact. Home shows adjusted vacancies; Excel includes adjusted vacancies and a Cases column listing each saved assignment/reason or pending review. The Summary sheet mirrors the Home page with side-by-side Admin & HR and Spinning tables. All worksheets scroll freely without frozen panes. Cases are not assigned to individual employees. Against Post is blue, Social Security is muted brown, and Other is neutral grey.

Select any designation on the Home Page to view the employees matched to its category, subcategory and designation. The responsive employee dialog includes all uploaded employee fields, search across those fields, and pagination. Configuration status is used only on the Configuration page and is omitted from the Home Page and all Excel sheets.

## Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

You can start editing the page by modifying `app/page.tsx`. The page auto-updates as you edit the file.

This project uses [`next/font`](https://nextjs.org/docs/app/building-your-application/optimizing/fonts) to automatically optimize and load [Geist](https://vercel.com/font), a new font family for Vercel.

## Learn More

To learn more about Next.js, take a look at the following resources:

- [Next.js Documentation](https://nextjs.org/docs) - learn about Next.js features and API.
- [Learn Next.js](https://nextjs.org/learn) - an interactive Next.js tutorial.

You can check out [the Next.js GitHub repository](https://github.com/vercel/next.js) - your feedback and contributions are welcome!

## Deploy on Vercel

The easiest way to deploy your Next.js app is to use the [Vercel Platform](https://vercel.com/new?utm_medium=default-template&filter=next.js&utm_source=create-next-app&utm_campaign=create-next-app-readme) from the creators of Next.js.

Check out our [Next.js deployment documentation](https://nextjs.org/docs/app/building-your-application/deploying) for more details.
