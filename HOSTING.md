# Hosting on GitHub Pages

`npm run dev` serves the panel from your Mac at `https://localhost:3000`, so only that Mac can load it. A Windows PC, or Excel on the web, needs the panel at an HTTPS address it can reach. GitHub Pages hosts it for free.

**Published.** The panel is live at https://bgage72590.github.io/excel-coach/taskpane.html, and the Windows manifest is at https://bgage72590.github.io/excel-coach/manifest.xml. Skip to [3. Install on Windows](#3-install-on-windows).

**Updating it:** commit your changes, then run `npm run publish:pages`. It commits this project minus the research and planning notes to a local `public` branch under your GitHub no-reply address, and pushes it to github.com/bgage72590/excel-coach, whose workflow tests, builds and deploys the site in about two minutes. Your own history and email never leave the Mac. Steps 1 and 2 below describe the one-time setup that's already done.

## What you get

- **A second add-in, Excel Coach (web).** It has its own add-in Id, so it installs next to the local Excel Coach instead of replacing it, and its ribbon button says Excel Coach (web), so you can tell the two apart. On your Mac, keep using the local one.
- **Only code on the site.** The site holds the built panel, its icons and its manifest. The coach generates practice data inside Excel each time it sets up a sheet, and your workbooks never leave Excel.
- **No AI hints.** **Explain my mistake** needs the local dev server, which holds your API key. Pages has no server, so the button doesn't appear. Everything else works the same.
- **Progress per computer.** Progress is saved in the panel's browser storage, so the Windows copy starts fresh and keeps its own. In Excel on the web, it's kept per browser. Clearing the Office cache also clears Excel Coach (web)'s progress on that computer, and clearing a browser's cache and site data can clear it in Excel on the web. There's no way to export progress yet.

## How the build works

```sh
PAGES_URL=https://<user>.github.io/excel-coach/ npm run build:pages
npm run validate:pages
```

`build:pages` typechecks, builds the site into `dist/` under the address's path (`/excel-coach/`), and writes `dist/manifest.xml`. That's `manifest.xml` with three changes:

- Every `https://localhost:3000` URL moves to the same path under `PAGES_URL`: the panel, icons, `bt:Url` resources and support URL. The app domain becomes the Pages origin.
- The add-in Id becomes the fixed hosted Id, `c16d946b-e5c7-4067-bc7a-5340a1bd6aa5`.
- The display name and the ribbon button's label become Excel Coach (web).

The build stops with an error if any address on this computer is left in the manifest, if the built code would load from one, or if a page loads files from outside the base path. A user site (`https://<user>.github.io/`) or a custom domain builds at `/`.

`validate:pages` checks `dist/manifest.xml` with Microsoft's validator, which sends the manifest to Microsoft's validation service. The validator also applies Microsoft Marketplace rules, so it rejects a `<Version>` below 1.0 in `manifest.xml`. Installing the add-in yourself doesn't need that.

Never change the hosted Id (`HOSTED_ADDIN_ID` in `scripts/pagesManifest.mjs`). Installed copies are tied to it, and a new Id means removing and reinstalling the add-in on every computer.

The workflow, `.github/workflows/pages.yml`, runs the same build on every push to `main`, or when you start it from the Actions tab. It runs `npm ci` and `npm test` first, takes `PAGES_URL` from the repository's Pages settings (`https://<owner>.github.io/<repo>/`, or your custom domain), and publishes `dist/`.

## 1. Create the repository

Free GitHub Pages needs a **public** repository. Only code is public there, never workbook data: the repository holds no workbooks, because the coach generates practice data at run time, and `.gitignore` keeps `.env.local` (your API key) and `dist/` out.

A private repository needs a paid plan (GitHub Pro, Team or Enterprise) to publish a site. Even then, the site itself is public unless your organization uses GitHub Enterprise Cloud. That's fine here, since the site holds only the panel.

Before you push, decide what else should be public. A public repository shows every tracked file and its full history:

- `research/` holds business, competitor and curriculum notes.
- `SCOPE.md`, `V2-PLAN.md` and `TESTING.md` are planning and test notes.
- Every commit shows its author's email address, and all of this project's commits use your Gmail address. To keep it private, on github.com go to **Settings › Emails**, turn on **Keep my email addresses private**, and copy the noreply address shown there. Then, in the project folder, run `git config user.email "<id>+<user>@users.noreply.github.com"` with that address. That covers new commits only.

These files and the address are already in this repository's commits, so a new commit doesn't keep them private: the push publishes the old commits too. To keep any of them private, push a fresh history without those files, with your noreply address on every commit. Claude can prepare one.

1. On github.com, select **+ › New repository**.
2. Name it `excel-coach`. The name becomes the site's path: `https://<user>.github.io/excel-coach/`.
3. Choose **Public**. Leave the README, .gitignore and license options off, because the project has its own.
4. Select **Create repository**.
5. In Terminal, in the project folder, push the `main` branch. The workflow publishes only from `main`.

   ```sh
   git remote add origin https://github.com/<user>/excel-coach.git
   git push -u origin main
   ```

The push starts the workflow before Pages is turned on, so that first run fails at the **Configure Pages** step. That's expected. Step 2 fixes it.

## 2. Turn on Pages and publish

1. In the repository, go to **Settings › Pages**.
2. Under **Build and deployment**, set **Source** to **GitHub Actions**. There's no branch or folder to pick.
3. Go to **Actions › Pages**, select **Run workflow**, keep the branch as `main`, and select **Run workflow**. Or open the failed run and select **Re-run all jobs**.
4. Wait for both jobs, **build** and **deploy**, to finish. It takes a few minutes. The deploy job shows the site's address. That address shows a 404 page, because the site has no home page. Add `taskpane.html` to the end.
5. Check the site in any browser:
   - `https://<user>.github.io/excel-coach/taskpane.html` shows the coach's home screen with "Open Excel Coach from the Home tab in Excel to set up practice sheets." Outside Excel, that's all it can do.
   - `https://<user>.github.io/excel-coach/manifest.xml` shows the manifest, with `Excel Coach (web)` and no `localhost`.

## 3. Install on Windows

Every install uses the hosted manifest. On the Windows PC, open `https://<user>.github.io/excel-coach/manifest.xml` in the browser and save it with Ctrl+S as `manifest.xml`. Don't use the `manifest.xml` in the project folder: that one loads the panel from your Mac.

### Excel for Windows: one PowerShell command (quickest)

This registers the hosted manifest as a developer add-in, the same way Microsoft's own sideloading tools do. Close Excel, open PowerShell (not as administrator), and run:

```powershell
$dir = "$env:USERPROFILE\ExcelCoach"; New-Item -ItemType Directory -Force $dir | Out-Null; Invoke-WebRequest "https://bgage72590.github.io/excel-coach/manifest.xml" -OutFile "$dir\manifest.xml"; $key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"; New-Item -Path $key -Force | Out-Null; New-ItemProperty -Path $key -Name "$dir\manifest.xml" -Value "$dir\manifest.xml" -PropertyType String -Force | Out-Null
```

Open Excel, then **Home › Add-ins**: **Excel Coach (web)** is listed under developer add-ins. To remove it, delete that value from the same registry key (or run `Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name "$env:USERPROFILE\ExcelCoach\manifest.xml"`). A work PC whose IT blocks custom add-ins won't show it; use one of the methods below or ask IT.

### Excel on the web

1. Go to microsoft365.com, open Excel, and open a workbook.
2. Select **Home › Add-ins**, then **More Settings**. (Some versions say **More Add-ins**.)
3. In the **Office Add-ins** dialog, select **Upload My Add-in**.
4. Select **Browse**, pick `manifest.xml`, and select **Upload**.

The **Coach** group appears on the Home tab. Select **Excel Coach (web)** to open the panel.

Excel keeps the uploaded manifest in the browser's storage. In another browser, or after you clear this browser's cache, upload it again. If **Upload My Add-in** is missing, a work or school account's administrator may have turned off custom add-ins.

### Excel for Windows: shared folder

Excel for Windows installs your own add-ins from a trusted catalog: a shared network folder it checks for manifests. You set it up once.

1. Create a folder, for example `C:\ExcelCoach`, and move `manifest.xml` into it.
2. Right-click the folder and select **Properties**. On the **Sharing** tab, select **Share**.
3. Your own account is listed as **Owner**, which is enough. Select **Share**. Copy the network path Windows shows, like `\\DESKTOP-ABC123\ExcelCoach`, then select **Done** and **Close**.
4. In Excel, select **File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs**.
5. Paste the network path into **Catalog Url** and select **Add catalog**. It must be the `\\computer\folder` path; Excel doesn't accept `C:\ExcelCoach`.
6. Select **Show in Menu** for the new catalog, then select **OK** twice.
7. Close Excel and open it again.
8. Select **Home › Add-ins**, then **Advanced**. In the **Office Add-ins** dialog, select **Shared Folder**, select **Excel Coach (web)**, and select **Add**.

The **Coach** group appears on the Home tab with the **Excel Coach (web)** button.

### Excel for Windows: Upload My Add-in

Some versions of Excel for Windows show **Upload My Add-in** in the same **Office Add-ins** dialog. If yours does, upload `manifest.xml` there, as on the web, and skip the shared folder.

### What works where

The panel adjusts its key names and menu paths for Windows and for Excel on the web. Excel on the web can't do everything the exercises teach: VBA macros don't run there, and Power Query and Solver are limited. Do those exercises in Excel for Windows.

## 4. Updating

- **Panel changes.** Push to `main`. The workflow publishes in a few minutes, and Excel loads the new panel the next time the task pane opens. To reload an open pane on Windows, select anywhere in it and press Ctrl+F5. GitHub Pages lets browsers cache pages for up to 10 minutes, so a new change can take that long to show. Panel changes never need a cache clear.
- **Manifest changes.** The ribbon button, name, icons and permissions come from the manifest, which Excel caches when you install it. After changing `manifest.xml`, raise its `<Version>`, push, and install the new hosted manifest on each computer. On the web, clear the browser's cache and upload it again. With the shared folder, replace the file in the folder, clear the Office cache (next item), and add it again from **Shared Folder**. Either cache clear can also clear Excel Coach (web)'s progress on that computer, so do it only for manifest changes.
- **Excel for Windows showing an old version.** First select anywhere in the pane and press Ctrl+F5, and allow 10 minutes after a push. If the ribbon or the panel is still old, clear the Office cache: close Excel, delete everything inside `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and open Excel again. Delete the folder's whole contents, not single files. Clearing the Office cache also clears Excel Coach (web)'s progress on that computer, and removes every add-in you installed yourself, so add Excel Coach (web) again from **Shared Folder**. Instead of deleting files, you can select **Next time Office starts, clear all previously-started web add-ins cache** in **Trusted Add-in Catalogs** and restart Excel. It clears progress the same way.
- **Renaming the repository** changes the address. The next build uses the new address, but installed copies still point at the old one, so install the new hosted manifest on each computer.

## 5. Removing

- **Excel on the web:** clear the browser's cache. The uploaded manifest lives there, so that removes the add-in. It can also clear the progress kept in that browser.
- **Excel for Windows:** in **File › Options › Trust Center › Trust Center Settings › Trusted Add-in Catalogs**, select the catalog, select **Remove**, then select **OK** twice. Then clear the Office cache as described in Updating, which removes the Coach group from the ribbon. Clearing the Office cache also clears Excel Coach (web)'s progress on that computer. To stop sharing the folder, right-click it, select **Properties › Sharing › Advanced Sharing**, and clear **Share this folder**.
- **The site:** in **Settings › Pages**, next to **Your site is live at**, open the **⋯** menu and select **Unpublish site**. Then go to **Actions › Pages**, open the **⋯** menu, and select **Disable workflow**, so the next push doesn't publish it again. Deleting the repository also takes the site down.

None of this affects the local Excel Coach on your Mac.

## When something goes wrong

| What you see | What to do |
|--------------|------------|
| `git push` fails with "GH007: Your push would publish a private email address" | Your GitHub account blocks pushes that show your private email address. Set git's `user.email` to your noreply address, as in step 1, and have Claude prepare a fresh history that uses it. Then push again. |
| The run fails at **Configure Pages** with "Get Pages site failed" | Pages isn't on yet. Do step 2, then run the workflow again. |
| The deploy job says the branch isn't allowed to deploy to `github-pages` | Only `main` publishes. Merge into `main` and push. |
| The build stops with "would still load from this computer" or "refers to this computer" | Something points at the dev server. The message lists each place. In `manifest.xml`, use `https://localhost:3000` (the build moves it) or remove the URL. In the panel's code, call the dev server with a relative path, like `/api/hint/status`. |
| The build stops with "doesn’t load its scripts from /excel-coach/assets/" or "outside /excel-coach/" | `vite.config.ts` must keep `base: process.env.COACH_BASE ?? '/'`. |
| The site's address shows a 404 page | That's expected. The site has no home page. Add `taskpane.html` to the end. |
| The task pane is blank | Open the `taskpane.html` address in a browser. If the home screen doesn't show there either, check the latest run's build log. |
| **Shared Folder** or **My Add-ins** doesn't list Excel Coach (web) | Check that you used the hosted `manifest.xml`, that it sits directly inside the shared folder, and that you restarted Excel after adding the catalog. |
| There's no **Explain my mistake** button | That's expected on Pages. AI hints need the local dev server. |

## Files

- `scripts/pagesManifest.mjs`, with types in `pagesManifest.d.mts`: the manifest transform, the hosted Id and name, the `PAGES_URL` checks, and the checks on the built site.
- `scripts/build-pages.mjs`: the build behind `npm run build:pages`.
- `.github/workflows/pages.yml`: test, build and publish.
- `tests/pages.test.ts`: every local URL moves, the Id differs and stays fixed, base paths are right, and leftovers fail the build, in the manifest or the built site.
