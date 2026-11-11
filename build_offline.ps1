$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TesseractDir = if ($env:TESSERACT_DIR) { $env:TESSERACT_DIR } else { Join-Path $ProjectDir 'tesseract-source' }
$DistDir = Join-Path $ProjectDir 'dist\PhanMemDoiTenPDF'
$ExePath = Join-Path $DistDir 'PhanMemDoiTenPDF.exe'

if (-not (Test-Path (Join-Path $TesseractDir 'libtesseract-5.dll'))) {
    throw "Không tìm thấy libtesseract-5.dll trong $TesseractDir. Hãy đặt biến môi trường TESSERACT_DIR tới thư mục Tesseract."
}
if (-not (Test-Path (Join-Path $TesseractDir 'tessdata\vie.traineddata'))) {
    throw "Không tìm thấy tessdata\vie.traineddata trong $TesseractDir."
}

Push-Location $ProjectDir
try {
    py -m PyInstaller --noconfirm --clean --onedir --windowed --name PhanMemDoiTenPDF app_gui.py
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }

    $TessOut = Join-Path $DistDir 'tesseract'
    New-Item -ItemType Directory -Force -Path $TessOut | Out-Null
    Copy-Item -Path (Join-Path $TesseractDir '*.dll') -Destination $TessOut -Force
    New-Item -ItemType Directory -Force -Path (Join-Path $TessOut 'tessdata') | Out-Null
    Copy-Item (Join-Path $TesseractDir 'tessdata\vie.traineddata') (Join-Path $TessOut 'tessdata') -Force

    $DataOut = Join-Path $DistDir 'data'
    New-Item -ItemType Directory -Force -Path $DataOut | Out-Null
    Copy-Item (Join-Path $ProjectDir 'data\Danh_muc_tai_lieu_khong_dau.xlsx') $DataOut -Force
    Copy-Item (Join-Path $ProjectDir 'README.md') $DistDir -Force
    Copy-Item (Join-Path $ProjectDir 'Chay_Phan_Mem_Doi_Ten.bat') $DistDir -Force

    Write-Host "Đã tạo gói offline tại: $DistDir"
    Write-Host "Khởi chạy: $ExePath"
} finally {
    Pop-Location
}
