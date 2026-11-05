function Assert-OpenDocsProductDirectory {
    param([string]$Root, [string]$ProductName)
    $resolvedRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $destination = [IO.Path]::GetFullPath((Join-Path $resolvedRoot $ProductName))
    if (-not $destination.StartsWith($resolvedRoot + '\', [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $destination -Leaf) -ne $ProductName) { throw 'Unsafe product directory.' }
    foreach ($path in @($resolvedRoot,$destination)) {
        if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Installation paths must not be reparse points.' }
    }
    return $destination
}
