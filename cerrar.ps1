# Cierra un .docx generado por ooxmlkit: actualiza los campos, incrusta las
# tipografias y exporta el PDF.
#
#     powershell -File cerrar.ps1 ejemplo\informe.docx
#
# El indice, el total de paginas y la incrustacion de fuentes no se pueden
# escribir desde python-docx: son campos que solo Word resuelve al abrir el
# archivo. Este paso exige Windows y Microsoft Word instalado. Sin el, el .docx
# es valido y editable, pero su indice sale vacio.

param([Parameter(Mandatory = $true)][string]$Docx)

$ErrorActionPreference = "Stop"

# El cast a [string] es necesario: Resolve-Path devuelve un PSObject y
# ExportAsFixedFormat rechaza ese tipo al pasarlo por [ref].
[string]$docx = (Resolve-Path $Docx).Path
[string]$pdf = [IO.Path]::ChangeExtension($docx, ".pdf")

# Si el PDF esta abierto en un visor, Word no puede sobrescribirlo y la
# exportacion falla a mitad. Se detecta aqui, con un mensaje claro en vez de
# una excepcion COM opaca.
if (Test-Path $pdf) {
    try {
        $fs = [IO.File]::Open($pdf, "Open", "ReadWrite", "None")
        $fs.Close()
    } catch {
        throw "El PDF esta abierto en otro programa. Cierralo y reintenta."
    }
}
$antes = if (Test-Path $pdf) { (Get-Item $pdf).LastWriteTime } else { [datetime]::MinValue }

$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0
try {
    $doc = $app.Documents.Open($docx)

    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }

    $doc.EmbedTrueTypeFonts = $true
    $doc.SaveSubsetFonts = $false
    $doc.DoNotEmbedSystemFonts = $true

    $doc.Repaginate()
    $paginas = $doc.ComputeStatistics(2)
    $doc.Save()

    # 17 = wdExportFormatPDF. Los $true finales activan los marcadores de
    # navegacion desde los titulos y las propiedades del documento.
    $doc.ExportAsFixedFormat([ref]$pdf, [ref]17, [ref]$false, [ref]0, [ref]0, [ref]0,
                             [ref]0, [ref]0, [ref]$true, [ref]$true, [ref]1,
                             [ref]$true, [ref]$true, [ref]$false)

    $doc.Close([ref]$false)
} finally {
    # Pase lo que pase, cerrar Word: si la exportacion lanza, sin este finally
    # queda un WINWORD huerfano reteniendo el .docx.
    $app.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}

# No confiar en que no hubo excepcion: comprobar que el PDF se reescribio de
# verdad. Sin esto, un fallo silencioso deja un PDF viejo con aspecto de nuevo.
if (-not (Test-Path $pdf) -or (Get-Item $pdf).LastWriteTime -le $antes) {
    throw "El PDF no se actualizo. Revisar si Word pudo exportarlo."
}

Write-Output "paginas: $paginas"
Write-Output $pdf
