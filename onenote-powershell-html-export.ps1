# References & sources used to create script:
  # https://passbe.com/2019/08/01/bulk-export-onenote-2013-2016-pages-as-html/
  # https://stackoverflow.com/questions/53689087/powershell-and-onenote
  # http://thebackend.info/powershell/2017/12/onenote-read-and-write-content-with-powershell/
  # https://stackoverflow.com/questions/53639041/how-to-access-contents-of-onenote-page

# Get export folder
Function Get-Folder($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
    $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
    $foldername.Description = "Select an export folder"
    $foldername.rootfolder = "MyComputer"
    if($foldername.ShowDialog() -eq "OK")
    {
        $folder += $foldername.SelectedPath
    }
    return $folder
}

# Spider and find each page, create directory for each group, and build TOC HTML
Function Spider-OneNote-Notebook {
    param( $onenote, $node, $path, $notebookRoot )
    $tocHtml = ""
    $previouslevel = 0
    $previousname = ""
    $grandparent = ""
    $parent = ""

    foreach($child in $node.ChildNodes) {
        $safeName = ReplaceIllegal -text $child.name
        $levelchange = $child.pageLevel - $previouslevel
        $displayName = [System.Net.WebUtility]::HtmlEncode($child.name)

        if (-not $child.HasChildNodes) {
            # --- It's a Page ---
            if ($levelchange -eq 1) {
                if ($previouslevel -ne 0) {
                    $grandparent = $parent
                    $parent = $previousname
                }
                $filepath = Join-Path -path $(join-path -path $path -ChildPath $grandparent) -ChildPath $parent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
            } elseif ($levelchange -eq -1) {
                $filepath = Join-Path -path $path -ChildPath $grandparent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                $parent = $grandparent
                $grandparent = ""
            } elseif ($levelchange -eq -2) {
                $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $path
                $parent = ""
                $grandparent = ""
            } elseif ($levelchange -eq 0 -and $parent -eq "") {
                $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $path
            } else {
                $grandparentpath = Join-Path -path $path -ChildPath $grandparent
                $filepath = Join-Path -path $grandparentpath -ChildPath $parent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                $fileAbsPath = Export-OneNote-Page -onenote $onenote -node $child -path $filepath
            }
            
            # Create a relative link for the index.htm
            if ($fileAbsPath) {
                $relPath = $fileAbsPath.Substring($notebookRoot.Length + 1)
                $relUrl = $relPath -replace '\\', '/' -replace ' ', '%20'
                
                # Use margin-left to visually indent subpages based on their OneNote level
                $indentLevel = [int]$child.pageLevel * 20
                $tocHtml += "<li class='page' style='margin-left: $($indentLevel)px;'><a href=`"$relUrl`">$displayName</a></li>`n"
            }

        } else {
            # --- It's a Section or Section Group ---
            $folder = Join-Path -Path $path -ChildPath $safeName
            New-Item -Path $folder -ItemType directory -ErrorAction Ignore | Out-Null
            Write-Host "  Section: $($folder)"

            $tocHtml += "<li class='section'>$displayName<ul>`n"
            # Recursively crawl the section and append its HTML
            $childToc = Spider-OneNote-Notebook -onenote $onenote -node $child -path $folder -notebookRoot $notebookRoot
            $tocHtml += $childToc
            $tocHtml += "</ul></li>`n"
        }

        # Store page level & name for next loop iteration
        $previouslevel = $child.pageLevel
        $previousname = $safeName 
    }
    
    return $tocHtml
}

# Export page
Function Export-OneNote-Page {
    param( $onenote, $node, $path )
    # Replace invalid file characters
    $name = ReplaceIllegal -text $node.name
    $file = $(Join-Path -Path $path -ChildPath "$($name).htm")
    Write-Host "    Page: $($file)"
    
    # Export
    $onenote.Publish($node.ID, $file, 7, "")
    
    # [BUGFIX] Changed $child.Name to $name here so attachments export correctly
	$attachmentpath = Join-Path -Path $path -ChildPath ($name + "_files")
    Export-OneNote-Attachments -onenote $onenote -node $node -path $attachmentpath

    # Return the absolute path so the Spider function can link to it
    return $file
}

# Copy embedded attachments
Function Export-OneNote-Attachments {
    param ( $onenote, $node, $path )
    $xml = ''
    $schema = @{one="http://schemas.microsoft.com/office/onenote/2013/onenote"}
    $onenote.GetPageContent($node.ID, [ref]$xml)
    $xml | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | foreach {
        $file = Join-Path -Path $path -ChildPath $_.Node.preferredName
        Write-Host "      Attachment: $($file)"
        Copy-Item $_.Node.pathCache -Destination $file
    }
}

Function ReplaceIllegal {
    param ( $text )
    $illegal = [string]::join('',([System.IO.Path]::GetInvalidFileNameChars())) -replace '\\\\','\\\\'
    $replaced = $text -replace "[$illegal]",'_'
    return $replaced
}

# ================= MAIN EXECUTION =================

# Get export folder
$folder = Get-Folder

if (-not $folder) {
    Write-Host "No folder selected. Exiting."
    exit
}

# Connect to OneNote COM API
$OneNote = New-Object -ComObject OneNote.Application
[xml]$Hierarchy = ""
$OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

# Loop over each notebook
foreach ($notebook in $Hierarchy.Notebooks.Notebook ) {
    $name = ReplaceIllegal -text $notebook.name
    $nf = Join-Path -Path $folder -ChildPath $name
    Write-Host "Notebook: $($nf)"
    New-Item -Path $nf -ItemType directory -ErrorAction Ignore | Out-Null
    
    # Kick off the spidering and capture the generated HTML list
    $tocBody = Spider-OneNote-Notebook -onenote $OneNote -node $notebook -path $nf -notebookRoot $nf

    # Wrap the list in a clean, styled HTML document
    $safeNotebookName = [System.Net.WebUtility]::HtmlEncode($notebook.name)
    $indexHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$safeNotebookName - Index</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 20px; background-color: #f3f2f1; }
        .container { max-width: 900px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #7719aa; border-bottom: 2px solid #7719aa; padding-bottom: 10px; }
        ul { list-style-type: none; padding-left: 20px; }
        li { margin: 6px 0; }
        .section { font-size: 1.2em; font-weight: bold; margin-top: 20px; color: #333; }
        .page { font-size: 1em; font-weight: normal; }
        a { text-decoration: none; color: #0078d4; }
        a:hover { text-decoration: underline; color: #004578; }
    </style>
</head>
<body>
    <div class="container">
        <h1>$safeNotebookName</h1>
        <ul>
            $tocBody
        </ul>
    </div>
</body>
</html>
"@

    # Save the index.htm at the root of the exported Notebook folder
    $indexPath = Join-Path -Path $nf -ChildPath "index.htm"
    Set-Content -Path $indexPath -Value $indexHtml -Encoding UTF8
    Write-Host "  -> Created Table of Contents: $($indexPath)"
}

# Cleanup filelist.xml files
Get-ChildItem -path $folder filelist.xml -Recurse | foreach { Remove-Item -Path $_.FullName }
Write-Host "Export Complete!"
