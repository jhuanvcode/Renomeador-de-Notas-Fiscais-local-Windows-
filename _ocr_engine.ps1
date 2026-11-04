param (
    [Parameter(Mandatory=$true)]
    [string]$ImagePath
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }

function Await-WinRT($asyncOp, $type) {
    $asTask = $asTaskGeneric.MakeGenericMethod($type)
    $netTask = $asTask.Invoke($null, @($asyncOp))
    return $netTask.GetAwaiter().GetResult()
}

[Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime] | Out-Null

$fullPath = (Resolve-Path $ImagePath).Path
$opFile = [Windows.Storage.StorageFile]::GetFileFromPathAsync($fullPath)
$file = Await-WinRT $opFile ([Windows.Storage.StorageFile])

$opStream = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
$stream = Await-WinRT $opStream ([Windows.Storage.Streams.IRandomAccessStream])

$opDecoder = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
$decoder = Await-WinRT $opDecoder ([Windows.Graphics.Imaging.BitmapDecoder])

$opBitmap = $decoder.GetSoftwareBitmapAsync()
$bitmap = Await-WinRT $opBitmap ([Windows.Graphics.Imaging.SoftwareBitmap])

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) {
    $lang = [Windows.Globalization.Language]::new("pt-BR")
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
}

$opResult = $engine.RecognizeAsync($bitmap)
$result = Await-WinRT $opResult ([Windows.Media.Ocr.OcrResult])

$words = @()
foreach ($line in $result.Lines) {
    foreach ($word in $line.Words) {
        $words += [PSCustomObject]@{
            Text = $word.Text
            X = [math]::Round($word.BoundingRect.X, 1)
            Y = [math]::Round($word.BoundingRect.Y, 1)
            W = [math]::Round($word.BoundingRect.Width, 1)
            H = [math]::Round($word.BoundingRect.Height, 1)
        }
    }
}

$words | ConvertTo-Json -Compress
