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

# Spider and find each page, create directory for each group
Function Spider-OneNote-Notebook {
    param( $onenote, $node, $path )
    $previouslevel = 0
    $previousname = ""
    $grandparent = ""
    $parent = ""
    foreach($child in $node.ChildNodes) {
        $child.name = ReplaceIllegal -text $child.name
        $levelchange = $child.pageLevel - $previouslevel
        if (-not $child.HasChildNodes) {
            if ($levelchange -eq 1) {
                if ($previouslevel -ne 0) {
                    $grandparent = $parent
                    $parent = $previousname
                }
                $filepath = Join-Path -path $(join-path -path $path -ChildPath $grandparent) -ChildPath $parent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                Export-OneNote-Page -onenote $onenote -node $child -path $filepath

            } elseif ($levelchange -eq -1) {
                $filepath = Join-Path -path $path -ChildPath $grandparent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                Export-OneNote-Page -onenote $onenote -node $child -path $filepath
                $parent = $grandparent
                $grandparent = ""
            } elseif ($levelchange -eq -2) {
                Export-OneNote-Page -onenote $onenote -node $child -path $path
                $parent = ""
                $grandparent = ""
            } elseif ($levelchange -eq 0 -and $parent -eq "") {
                Export-OneNote-Page -onenote $onenote -node $child -path $path
            } else {
                $grandparentpath = Join-Path -path $path -ChildPath $grandparent
                $filepath = Join-Path -path $grandparentpath -ChildPath $parent
                New-Item -Path $filepath -ItemType directory -ErrorAction Ignore | Out-Null
                Export-OneNote-Page -onenote $onenote -node $child -path $filepath
            }
        } else {
            $folder = Join-Path -Path $path -ChildPath $child.name
            New-Item -Path $folder -ItemType directory -ErrorAction Ignore | Out-Null
            Write-Host "  Section: $($folder)"
            Spider-OneNote-Notebook -onenote $onenote -node $child -path $folder
        }
        # Store page level
        $previouslevel = $child.pageLevel
        $previousname = $child.name
    }

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
    Export-OneNote-Attachments -onenote $onenote -node $node -path $path
}

# Copy embedded attachments
Function Export-OneNote-Attachments {
    param ( $onenote, $node, $path )
    $xml = ''
    $schema = @{one=”http://schemas.microsoft.com/office/onenote/2013/onenote”}
    $onenote.GetPageContent($node.ID, [ref]$xml)
    $xml | Select-Xml -XPath "//one:Page/one:Outline/one:OEChildren/one:OE/one:InsertedFile" -Namespace $schema | foreach {
        $file = Join-Path -Path $path -ChildPath $_.Node.preferredName
        Write-Host "      Attachment: $($file)"
        Copy-Item $_.Node.pathCache -Destination $file
    }
}

Function ReplaceIllegal {
    param ( $text )
    $illegal = [string]::join('',([System.IO.Path]::GetInvalidFileNameChars())) -replace '\\','\\'
    $replaced = $text -replace "[$illegal]",'_'
    return $replaced
}

# Get export folder
$folder = Get-Folder

# Connect
$OneNote = New-Object -ComObject OneNote.Application
[xml]$Hierarchy = ""
$OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

# Loop over each notebook
foreach ($notebook in $Hierarchy.Notebooks.Notebook ) {
    $name = ReplaceIllegal -text $notebook.name
    $nf = Join-Path -Path $folder -ChildPath $name
    Write-Host "Notebook: $($nf)"
    New-Item -Path $nf -ItemType directory | Out-Null
    Spider-OneNote-Notebook -onenote $OneNote -node $notebook -path $nf
}
