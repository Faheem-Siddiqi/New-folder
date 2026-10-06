Kohinoor workforce reporting app: upload a turnover workbook on `/`, review strength results, and download a formatted Excel report. `/config` manages the strength template.

## Report workflow

1. Upload one `.xlsx` or `.xls` workbook (up to 20 MB). Open or repair unreadable files in Excel and save a new, unencrypted copy.
2. Scanning detects all 12 turnover headers by name, in any order, after title rows. Complete tables on multiple worksheets and repeated headers are supported. Common header aliases, punctuation, case, whitespace, and explicit merged department/unit cells are handled. Incomplete sections and unreadable worksheets are flagged; irrecoverable workbook corruption produces an error.
3. Review category results and validation notes. On-roll means the number of unique listed employees that match a template designation; the upload must represent the workforce you want to count. The provided columns contain no termination/status field, so the app does not infer departures or turnover rates.
4. Download a genuine `.xlsx` workbook with Summary, Strength Detail, Employees, Validation, and one sheet per configured category. Headers, filters, frozen rows, bounded column widths, formula-based vacancies, and the official logo are included.

Matching uses normalized department/unit/subcategory context, designation, and grade. Department + Unit can match a combined subcategory such as `SPINNING - RING - 1`. Category/Org Group can disambiguate candidates. Ambiguous and unmatched records are excluded from counts and listed for review. Employee number + Unit identifies duplicates; identical records count once and conflicting identities are excluded.

`strength-master-seed-0.json` is the existing app template. Scanning and exports only read its data; only the configuration save action writes it. Designation, grade, cadre, approved strength, and status come from the template. It contains no individual employee names, IDs, gender, shifts, or joining dates: those appear in the Employees sheet from the upload, with separate matched template columns.

Successful reports and original workbook blobs are stored in IndexedDB on the current browser/device. Reloading, reopening the page, and navigation from Configuration restore the saved report. A changed template triggers reprocessing of the saved workbook. Failed or cancelled scans retain the previous report. Storage denial/quota failures are reported and the current report remains downloadable; private browsing or clearing site data can remove saved reports. Uploads are processed locally in Web Workers and are not sent to an external service.

Logo source: [Kohinoor Textile Mills official website](https://www.kmlg.com/ktml/wp-content/uploads/2016/05/logo-black.png), saved locally as `public/kohinoor-logo.png`.

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
