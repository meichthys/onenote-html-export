# Onenote Html Export
Exports local onenote notebooks in html directory format.

The bulk of the code was taken from [@passbe](https://github.com/passbe)'s [blog post](https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/).
The main change in this repository is the handling of onenote sub-pages and sub-sub-pages. If you would prefer sub-pages and sub-sub-pages to be exported in the same directory use [@passbe](https://github.com/passbe)'s script.

The output of [onenote-powershell-html-export.ps1](onenote-powershell-html-export.ps1) will look like the following. The script will export all notebooks, sections, pages, sub-pages, sub-sub-pages, and attachments:

```
Export Directory
└── Notebook
    └── Section
        ├── page.htm
        ├── Page
            ├── sub-page.htm
            └── Subpage
                └── sub-sub-page.htm
        └── Page_files
            └── Page_files
                └── file.png
...
```
