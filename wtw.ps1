param (
    [string]$mhtFilePath = $args[0]
)

try {
    # Resolve relative path to absolute path
    $mhtFilePath = (Resolve-Path -Path $mhtFilePath).Path

    # Generate the output file path by replacing the .mht extension with .html
    $htmlFilePath = [System.IO.Path]::ChangeExtension($mhtFilePath, ".html")

    # Create a new Word application object
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    # Open the .mht file
    $document = $word.Documents.Open($mhtFilePath)

    # Save the document as .html
    $document.SaveAs([ref] $htmlFilePath, [ref] 8)  # 8 corresponds to wdFormatHTML

    # Close the document and quit Word
    $document.Close()
    $word.Quit()

    # Release the COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

    # Garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Output "Conversion complete: $htmlFilePath"
} catch {
    Write-Error "An error occurred: $_"
    if ($word) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($document) {
        $document.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
