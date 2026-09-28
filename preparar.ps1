#Requires -Version 5.1
<#
.SYNOPSIS
    Genera el manifiesto del complemento para la URL donde se publica la web.
.EXAMPLE
    .\preparar.ps1 -Url https://jnadalcdi.github.io/imprimir-correo-outlook   (GitHub Pages)
    .\preparar.ps1 -Url https://localhost:3000                                (pruebas locales)
.NOTES
    Resultado: manifest-final.xml.  Ese fichero es el que se sube al centro
    de administración de Microsoft 365 (o se carga en Outlook para probar).
    El contenido del repositorio es lo que se publica en esa URL.
#>
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://')]
    [string]$Url
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$Url = $Url.TrimEnd('/')
$src = Join-Path $PSScriptRoot "manifest.xml"
$dst = Join-Path $PSScriptRoot "manifest-final.xml"

# AppDomain solo admite el dominio (p. ej. GitHub Pages: https://usuario.github.io).
$origin = ([Uri]$Url).GetLeftPart([UriPartial]::Authority)
$xml = [IO.File]::ReadAllText($src).Replace("__BASE_URL__", $Url).Replace("__BASE_ORIGIN__", $origin)
[IO.File]::WriteAllText($dst, $xml, (New-Object Text.UTF8Encoding $false))

Write-Host ""
Write-Host "  Manifiesto generado: $dst" -ForegroundColor Green
Write-Host "  URL de la web:       $Url/  (publicar ahi el contenido del repositorio)" -ForegroundColor White
Write-Host ""
