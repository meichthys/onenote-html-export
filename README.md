# Onenote Html Export
Exports local onenote notebooks in html directory format.

| :exclamation:  Note: If you want to convert an online notebook, import it into your desktop version of Onenote first. |
|-----------------------------------------------------------------------------------------------------------------------|

The bulk of the code was taken from [@passbe](https://github.com/passbe)'s [blog post](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/).
The main change in this repository is the handling of onenote sub-pages and sub-sub-pages. If you would prefer sub-pages and sub-sub-pages to be exported in the same directory use [@passbe](https://github.com/passbe)'s [blog post](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/) script.

The output of [onenote-powershell-html-export.ps1](onenote-powershell-html-export.ps1) will look like the following. The script will export all notebooks, sections, pages, sub-pages, sub-sub-pages, and attachments:

```
Export Directory
└── notebook
    └── section
        ├── page.htm
        ├── page
            ├── sub-page.htm
            └── sub-page
                └── sub-sub-page.htm
        └── page_files
            └── file.png
...
```

## Usage
- Have powershell installed
- Open Onenote Desktop (Only Microsoft Office Versions of Onenote are supported by this script)
  - Import notebooks if exported from OneNote web version
- Run the powershell script: `onenote-powershell-html-export.ps1`
  - Select your export directory.
  - Watch yourself exit Microsoft's walled garden
  - Check console for any errors (most can be ignored: see below)

## Notes
- Errors that look like the following can generally be ignored. They are usually a result of a deleted page or an outlook attachment/task: ![image](https://user-images.githubusercontent.com/10717998/132024369-ec7a51e4-be6b-4d7e-855f-d803126c4eda.png)
- Page names containing invalid file/directory charachters are replaced with an underscore: `10_00am.htm` replaces `10:00am`
- Currently PDF attachments are not being saved to the `_files` subdirectory, but rather too the same directory of the page they were attached to.
