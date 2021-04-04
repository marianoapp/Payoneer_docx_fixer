function arraysMatch {
    param (
        $firstArray,
        $secondArray,
        $startIndex
    )

    for ($i=0; $i -lt $secondArray.length; $i++) {
        if ($firstArray[$startIndex+$i] -ne $secondArray[$i]) {
            return $false
        }
    }
    return $true
}

# validate that an argument was provided
if ($args.count -ne 1) {
    write-host "Usage: fix_docx.ps1 file"
    exit(1)
}

$filePath = $args[0]
$file = get-item $filePath
$workFolder = $file.directoryName
$newFileBasename = "$($file.basename)_fixed.docx"
$newFile = "$workFolder\$newFileBasename"

# check if a file with the same name as the fixed one already exists
if (test-path $newFile) {
    write-host "A file named '$newFileBasename' already exists, aborting"
    exit(1)
}

# check the file extension
if ($file.extension -ne ".docx") {
    write-host "Wrong file type, it should be a Word document (.docx)"
    exit(1)
}

# expected file header
# 4 bytes magic number
# 2 bytes minimum version to decompress (we don't care)
# 2 bytes general purpose bit flags
$headerSize = 8

# read the signature bytes
$fs = [IO.File]::OpenRead($file)
$buffer = [byte[]]::new($headerSize)
$fs.Read($buffer, 0, $buffer.length) | out-null
$fs.Close()

# validate the header (magic number)
$magicNumber = @(0x50, 0x4b, 0x03, 0x04)
if (!(arraysMatch $buffer $magicNumber 0)) {
    write-host "This is not a valid file"
    exit(1)
}

# validate the flags
$fixedFlags = @(0x06, 0x00)
if (arraysMatch $buffer $fixedFlags 6) {
    write-host "This file is already valid"
    exit(0)
}

# create the fixed file
copy-item $file $newFile

# write the correct flags to the header
$fs = [IO.File]::OpenWrite((get-item $newFile))
$fs.Seek(6, "Begin") | out-null
$fs.WriteByte(0x06)
$fs.WriteByte(0x00)
$fs.Flush()
$fs.Close()

write-host "File fixed"