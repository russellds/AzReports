function CheckPath {
    param(
        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $Path,

        [string]
        $Extension,

        [switch]
        $Force
    )

    if ($Extension) {
        if ($Path.Extension -ne $Extension) {
            throw "File extension must be $( $Extension )!"
        }
    }

    if (Test-Path -Path $Path.DirectoryName) {
        if (Test-Path -Path $Path.FullName) {
            if ($Force) {
                [void](Remove-Item -Path $Path.FullName -Force)
            } else {
                throw "$( $Path.FullName ) already exists, pass -Force to overwrite!"
            }
        }
    } else {
        [void](New-Item -Path $Path.DirectoryName -ItemType Directory -Force)
    }
}
