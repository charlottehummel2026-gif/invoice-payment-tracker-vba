# Put the project on GitHub

Project repository: [charlottehummel2026-gif/invoice-payment-tracker-vba](https://github.com/charlottehummel2026-gif/invoice-payment-tracker-vba).

Portfolio demo: [Invoice & Payment Tracker](https://charlottehummel2026-gif.github.io/invoice-payment-tracker-vba/).

The instructions below also apply when publishing a new copy of the project.

Suggested repository name: `invoice-payment-tracker-vba`.

Suggested description: **Excel/VBA receivables tracker with validated entry forms, payment matching, aging reports, and an interactive portfolio demo.**

## Upload the package

1. Create a new GitHub repository. Use a public repository for a publicly reviewable resume portfolio.
2. Upload the contents of this portfolio folder, keeping its folder structure. `index.html`, `README.md`, and `package.json` should sit at the repository root.
3. Include `assets`, `downloads`, `src`, `docs`, `tests`, `.gitignore`, and `.nojekyll`. Upload the extracted files, not only the ZIP archive.
4. Commit the files. The downloadable workbook is already part of `downloads/`.

The package contains only the fictional invoice tracker and its portfolio files. The original Funds Tracker template and local working folders are not needed in the repository.

## Enable the website

In the repository, go to **Settings > Pages**. Under **Build and deployment**, choose **Deploy from a branch**, then select **main** and **/(root)**. Save. Use the **Visit site** link once GitHub reports that the site is published. See [GitHub's publishing-source instructions](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

The site is static and requires no API keys, package installation, or server-side setup. GitHub Pages uses the root `index.html`; `.nojekyll` keeps it as a plain static site. [GitHub Pages setup reference](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site).

## Connect it to your resume

- Pin the repository on your GitHub profile.
- Add the published Pages URL to the repository's **About > Website** field.
- Put the live demo and repository links next to the project title on your resume.
- Add the live URL near the top of `README.md` after publishing.

Suggested resume entry:

**Invoice & Payment Tracker | Excel, VBA, Data Validation, Reporting**  
Developed an Excel/VBA receivables tracker with validated entry forms, automatic IDs, payment allocation, aging analysis, and dashboard reporting.

## Preview before publishing

Open `index.html` directly in a browser. For an HTTP preview, run this from the portfolio folder and visit the printed address:

```sh
python3 -m http.server 8000
```

Run `npm test` to verify the browser calculation rules before uploading changes. Native Excel verification is tracked separately in `TESTING.md`.
