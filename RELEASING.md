# Releasing a new version

Notes for the developer. Users only need the README.

## 1. Update the add-in

The `.bas` / `.frm` files are the source; `ClientExportTool.xlam` is the built add-in. Keep them in sync.

**Source changed (e.g. you edited a `.bas` in this repo) → rebuild the `.xlam`:**

1. Open `ClientExportTool.xlam` in Excel (File → Open). Nothing visible opens; that's normal for an add-in.
2. Press **Alt+F11** (Mac: **Tools → Macro → Visual Basic Editor**).
3. In the Project pane, under **VBAProject (ClientExportTool.xlam)**, right-click the module you changed → **Remove…** → **No** (don't export).
4. **File → Import File…** → pick the changed `.bas` / `.frm` from this repo.
5. **Debug → Compile VBAProject**. Fix anything it flags.
6. Click **Save** in the VBA editor, then quit Excel.

**Code changed in the VBA editor → export it back:** right-click each changed module → **Export File…** over the copy in this repo.

Then commit both the source and the `.xlam`.

## 2. Publish

1. Push to `main`.
2. On GitHub: **Releases → Draft a new release**, create a tag like `v1.2`, write the notes, **Publish**.
3. The **Release packages** workflow (Actions tab) builds and attaches, within a minute or two:
   - `ClientExportTool-v1.2-Windows.zip` (add-in, `Install-Windows.cmd`, README)
   - `ClientExportTool-v1.2-macOS.zip` (add-in, `install-mac.command`, README)

To rebuild the zips for an existing release: **Actions → Release packages → Run workflow** and enter the tag.

Pull requests also run the workflow as a check; the zips it builds can be downloaded from the run page.

## Things that break the installers

- **Line endings.** `Install-Windows.cmd` must use CRLF and `install-mac.command` must use LF. `.gitattributes` enforces this on checkout, and the workflow refuses to package files with the wrong endings.
- **Non-ASCII characters** (smart quotes, em dashes, arrows) in `Install-Windows.cmd`. Keep it plain ASCII.
- **Renaming files.** The installers look for `ClientExportTool*.xlam` next to themselves and always install it as `ClientExportTool.xlam`.
