# FormaScholar

FormaScholar is a Microsoft Word on the web add-in for checking academic manuscripts against the AIP Conference Proceedings - September 2023 formatting requirements.

It reads the document that is open in Word, reports formatting findings, lets users locate issues, and offers only explicitly approved deterministic corrections. It does not upload documents to a server, change academic content, or require a database, Python service, or LLM.

## What it provides

- Automated checks for document layout, typography, headings, captions, equations, references, figures, tables, and related structure.
- A compliance score for applicable checks that can be evaluated automatically.
- Manual-review and submission-checklist findings for requirements that need human judgement or a final PDF.
- Safe corrections followed by automatic revalidation.
- Template analysis for creating a separate imported formatting strategy.

## Prerequisites

- Node.js current LTS and npm.
- A Microsoft 365 account with access to Word on the web.
- A supported web browser.

VS Code is optional. The included Windows scripts can install dependencies and start the project without it. Word desktop is not required for the target workflow.

## Installation

Clone the repository and install the npm dependencies:

```bash
git clone https://github.com/chow0226/FormaScholar.git
cd FormaScholar
npm install
```

`npm install` installs the dependency versions recorded in `package-lock.json`.

## Dependencies

### Runtime dependency

- `jszip` - reads the compressed DOCX package so the add-in can inspect WordprocessingML and formatting properties in the browser.

Office.js is supplied by the Word host. The project includes Office.js and Office Runtime type definitions for development.

### Development dependencies

- TypeScript and `ts-loader` for compiling the source code.
- Webpack and `webpack-dev-server` for bundling and serving the add-in.
- `office-addin-dev-certs` for the local HTTPS development certificate.
- `office-addin-manifest` for validating the XML manifest.
- CSS and HTML webpack plugins/loaders for packaging the task pane.

## Run locally

### Windows shortcut

Double-click `START_DEV.bat`. It will install dependencies if needed and start the HTTPS development server at:

```text
https://localhost:3000
```

### Command line

```bash
npm run dev
```

Keep the terminal running while using the add-in. The local server must use HTTPS because Office web add-ins require a secure origin.

## Open the add-in in Word on the web

1. Start the local server with `START_DEV.bat` or `npm run dev`.
2. Open a `.docx` manuscript in Word on the web.
3. Open **Home > Add-ins** and choose the advanced or additional settings option.
4. Select **Upload My Add-in**.
5. Choose the project's `manifest.xml` file and upload it.
6. Open the FormaScholar tab or button in the Word ribbon.
7. Choose **Open FormaScholar**.
8. Select **Validate current document**.

Microsoft may use slightly different menu labels depending on the Microsoft 365 tenant. The required action is uploading the local XML manifest through Word on the web.

## Typical workflow

1. Open a manuscript in Word on the web.
2. Open FormaScholar.
3. Select the AIP Conference Proceedings strategy.
4. Validate the current document.
5. Review the score and findings.
6. Use **Locate** to jump to an issue in Word.
7. Select supported corrections when appropriate.
8. Apply the approved corrections.
9. Review the automatic revalidation results.

Corrections are deliberately conservative. The add-in does not blindly reformat an entire manuscript or automatically alter academic meaning.

## Available commands

Run these commands from the project directory:

```bash
npm run dev               # Start the local HTTPS development server
npm run build             # Create the production bundle in dist/
npm run validate-manifest # Validate manifest.xml
```

Windows shortcuts are also available:

```text
START_DEV.bat             Start local development
BUILD.bat                 Create a production build
VALIDATE_MANIFEST.bat     Validate the XML manifest
```

The build output is generated in `dist/`. It is a generated folder and is not committed to Git.

## Project structure

```text
FormaScholar/
|-- manifest.xml                         Word add-in manifest for local use
|-- manifest.production.template.xml     Template for hosted deployment
|-- package.json                         Scripts and dependencies
|-- package-lock.json                    Locked npm dependency tree
|-- START_DEV.bat                        Windows development launcher
|-- BUILD.bat                            Windows build script
|-- VALIDATE_MANIFEST.bat                Windows manifest validation script
|-- assets/                              Add-in icons
|-- reference/                           Derived reference metadata
|-- src/
|   |-- core/                            Validation and correction logic
|   |-- ooxml/                           DOCX and OOXML inspection
|   |-- strategies/                      Built-in formatting strategies
|   |-- taskpane/                        Task pane UI and interaction
|   |-- template/                        Template analysis
|   `-- word/                            Word document adapter
```

The supplied reference DOCX is kept local and is not included in the public repository. The derived metadata file in `reference/` is included for development and audit purposes.

## Production deployment

The add-in has no server-side backend, but its compiled web files must be hosted on an HTTPS web host for production use.

1. Run `npm run build`.
2. Host the contents of `dist/` on an HTTPS static host.
3. Copy `manifest.production.template.xml` to a production manifest file.
4. Replace each `https://YOUR-HOST.example.com` value with the actual HTTPS origin.
5. Deploy the production manifest through the appropriate Microsoft 365 add-in distribution method.

End users do not need Node.js or a command-line environment after the compiled assets are hosted.

## Troubleshooting

### `node` or `npm` is not recognized

Install the current Node.js LTS release, reopen PowerShell, and run `node --version` and `npm --version`.

### Word cannot connect to the local add-in

Confirm that `npm run dev` is still running and that the add-in is using `https://localhost:3000`. Local Office add-ins require HTTPS.

### The Upload My Add-in option is unavailable

Some Microsoft 365 organizations restrict custom add-in sideloading. Contact the tenant administrator or use an account that permits development add-ins.

### The manifest is rejected

Run:

```bash
npm run validate-manifest
```

Then confirm that `manifest.xml` is the local development manifest and that the development server is running on the URL specified in that manifest.

## Limitations

- The built-in strategy focuses on AIP Conference Proceedings - September 2023.
- Visual judgement, semantic author-affiliation checks, image quality, final-PDF checks, and copyright evidence may require manual review.
- Unusual or severely malformed documents may produce conservative manual-review or not-evaluated findings.
- Imported strategies are stored locally in the browser and are not synchronized between users or devices.
