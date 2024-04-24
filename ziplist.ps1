[cmdletbinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$ZipFileName,
    [String]$ExportCSVFileName
)

# Resolve relative file paths to absolute paths
$ZipFileName = (Resolve-Path $ZipFileName).Path
if ($ExportCSVFileName) {
    $ExportCSVFileName = (Resolve-Path $ExportCSVFileName).Path
}

# Exit if the shell is using lower version of dotnet
$dotnetversion = [Environment]::Version
if(!($dotnetversion.Major -ge 4 -and $dotnetversion.Build -ge 30319)) {
    write-error "Microsoft DotNet Framework 4.5 is not installed. Script exiting"
    exit(1)
}

# Import dotnet libraries
[Void][Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

if(Test-Path $ZipFileName) {
    $RawFiles = [IO.Compression.ZipFile]::OpenRead($ZipFileName).Entries
    $ObjArray = foreach($RawFile in $RawFiles) {
        $object = New-Object -TypeName PSObject
        $Object | Add-Member -MemberType NoteProperty -Name FileName -Value $RawFile.Name
        $Object | Add-Member -MemberType NoteProperty -Name FullPath -Value $RawFile.FullName
        $Object | Add-Member -MemberType NoteProperty -Name CompressedLengthInKB -Value ($RawFile.CompressedLength/1KB).Tostring("00")
        $Object | Add-Member -MemberType NoteProperty -Name UnCompressedLengthInKB -Value ($RawFile.Length/1KB).Tostring("00")
        $Object | Add-Member -MemberType NoteProperty -Name FileExtn -Value ([System.IO.Path]::GetExtension($RawFile.FullName))
        $Object | Add-Member -MemberType NoteProperty -Name ZipFileName -Value $ZipFileName
        $Object
    }
} else {
    Write-Warning "$ZipFileName File path not found"
    exit(1)
}

if ($ExportCSVFileName){
    try {
        $ObjArray | Export-CSV -Path $ExportCSVFileName -NotypeInformation
    } catch {
        Write-Error "Failed to export the output to CSV. Details : $_"
    }
} else {
    $ObjArray | Format-Table -AutoSize
}