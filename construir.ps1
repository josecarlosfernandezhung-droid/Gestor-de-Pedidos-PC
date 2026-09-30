# Construye dist\ServidorPedidos.exe (un solo archivo, con el servidor cifrado adentro)
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $dir
# --- Icono: si icono.ico no existe o no es un .ico real (p.ej. un JPG/PNG renombrado), se genera desde el PNG/JPG ---
function Test-Ico([string]$p) {
    if (-not (Test-Path $p)) { return $false }
    $b = [IO.File]::ReadAllBytes($p)
    return ($b.Length -gt 22 -and $b[0] -eq 0 -and $b[1] -eq 0 -and $b[2] -eq 1 -and $b[3] -eq 0)
}
function New-Ico([string]$src, [string]$dst) {
    Add-Type -AssemblyName System.Drawing
    $img = [Drawing.Image]::FromFile($src)
    $sizes = 16, 24, 32, 48, 64, 128, 256
    $entries = @()
    foreach ($sz in $sizes) {
        $bmp = New-Object Drawing.Bitmap $sz, $sz, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [Drawing.Graphics]::FromImage($bmp)
        $g.Clear([Drawing.Color]::Transparent)
        $g.InterpolationMode = 'HighQualityBicubic'; $g.SmoothingMode = 'HighQuality'; $g.PixelOffsetMode = 'HighQuality'
        $g.DrawImage($img, 0, 0, $sz, $sz); $g.Dispose()
        # DIB 32 bpp (BGRA, filas de abajo hacia arriba) + mascara AND vacia
        $ms = New-Object IO.MemoryStream
        $bw = New-Object IO.BinaryWriter $ms
        $bw.Write([int]40); $bw.Write([int]$sz); $bw.Write([int]($sz * 2))
        $bw.Write([int16]1); $bw.Write([int16]32); $bw.Write([int]0)
        $bw.Write([int]($sz * $sz * 4)); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0)
        for ($y = $sz - 1; $y -ge 0; $y--) { for ($x = 0; $x -lt $sz; $x++) {
            $c = $bmp.GetPixel($x, $y); $bw.Write([byte]$c.B); $bw.Write([byte]$c.G); $bw.Write([byte]$c.R); $bw.Write([byte]$c.A) } }
        $maskRow = [int]([Math]::Ceiling($sz / 32.0) * 4)
        $bw.Write((New-Object byte[] ($maskRow * $sz)))
        $bw.Flush(); $entries += , @($sz, $ms.ToArray()); $bmp.Dispose()
    }
    $img.Dispose()
    $out = New-Object IO.MemoryStream
    $w = New-Object IO.BinaryWriter $out
    $w.Write([int16]0); $w.Write([int16]1); $w.Write([int16]$entries.Count)
    $offset = 6 + 16 * $entries.Count
    foreach ($e in $entries) {
        $sz = $e[0]; $data = $e[1]
        $dim = if ($sz -ge 256) { 0 } else { $sz }
        $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
        $w.Write([int16]1); $w.Write([int16]32); $w.Write([int]$data.Length); $w.Write([int]$offset)
        $offset += $data.Length
    }
    foreach ($e in $entries) { $w.Write([byte[]]$e[1]) }
    $w.Flush()
    [IO.File]::WriteAllBytes($dst, $out.ToArray())
}
$icoPath = Join-Path $dir 'icono.ico'
if (-not (Test-Ico $icoPath)) {
    $srcImg = @('icono_pc.png', 'icono.png') | ForEach-Object { Join-Path $dir $_ } | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $srcImg) { throw 'icono.ico no es valido y no hay icono_pc.png / icono.png para generarlo.' }
    Write-Host "Generando icono.ico desde $srcImg"
    New-Ico $srcImg $icoPath
}

foreach ($f in 'clave_publica.txt','servidor_pedidos.ps1','xlsx_full_min.js','Launcher.cs','app.manifest','icono.ico') {
    if (-not (Test-Path (Join-Path $dir $f))) { throw "Falta el archivo: $f" }
}
$pub = [Convert]::FromBase64String((Get-Content (Join-Path $dir 'clave_publica.txt') -Raw).Trim())
if ($pub.Length -ne 65 -or $pub[0] -ne 4) { throw 'clave_publica.txt no es una clave publica valida del generador.' }

$utf8 = New-Object System.Text.UTF8Encoding($false)
$ps = $utf8.GetBytes([IO.File]::ReadAllText((Join-Path $dir 'servidor_pedidos.ps1')))
$js = $utf8.GetBytes([IO.File]::ReadAllText((Join-Path $dir 'xlsx_full_min.js')))
$ms = New-Object IO.MemoryStream
$ms.Write([BitConverter]::GetBytes([int]$ps.Length), 0, 4); $ms.Write($ps, 0, $ps.Length); $ms.Write($js, 0, $js.Length)
$plain = $ms.ToArray()

$aes = [Security.Cryptography.Aes]::Create()
$aes.KeySize = 256; $aes.Mode = 'CBC'; $aes.Padding = 'PKCS7'; $aes.GenerateKey(); $aes.GenerateIV()
$enc = $aes.CreateEncryptor().TransformFinalBlock($plain, 0, $plain.Length)
$ms2 = New-Object IO.MemoryStream
$ms2.Write($aes.IV, 0, 16); $ms2.Write($enc, 0, $enc.Length)
[IO.File]::WriteAllBytes((Join-Path $dir 'payload.bin'), $ms2.ToArray())

$rng = [Security.Cryptography.RandomNumberGenerator]::Create()
function Mask([byte[]]$secret) {
    $m = New-Object byte[] $secret.Length; $rng.GetBytes($m)
    $x = New-Object byte[] $secret.Length
    for ($i = 0; $i -lt $secret.Length; $i++) { $x[$i] = $secret[$i] -bxor $m[$i] }
    return @(('new byte[]{' + ($m -join ',') + '}'), ('new byte[]{' + ($x -join ',') + '}'))
}
$hm = New-Object byte[] 32; $rng.GetBytes($hm)
$k = Mask $aes.Key; $h = Mask $hm
$cs = @"
namespace TotoLic { static class Secrets {
  static readonly byte[] A = $($k[0]); static readonly byte[] B = $($k[1]);
  static readonly byte[] C = $($h[0]); static readonly byte[] D = $($h[1]);
  public static readonly byte[] Pub = new byte[]{$($pub -join ',')};
  static byte[] X(byte[] a, byte[] b) { byte[] r = new byte[a.Length]; for (int i = 0; i < r.Length; i++) r[i] = (byte)(a[i] ^ b[i]); return r; }
  public static byte[] Key() { return X(A, B); }
  public static byte[] Hm() { return X(C, D); }
} }
"@
[IO.File]::WriteAllText((Join-Path $dir 'Secrets.g.cs'), $cs)

$csc = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path $csc)) { $csc = "$env:windir\Microsoft.NET\Framework\v4.0.30319\csc.exe" }
$sma = Get-ChildItem "$env:windir\Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation" -Recurse -Filter System.Management.Automation.dll | Select-Object -First 1 -ExpandProperty FullName
if (-not $sma) { throw 'No se encontro System.Management.Automation (PowerShell 5.1).' }
New-Item -ItemType Directory -Force -Path (Join-Path $dir 'dist') | Out-Null
$out = Join-Path $dir 'dist\ServidorPedidos.exe'
& $csc /nologo /target:exe /platform:anycpu /optimize+ "/out:$out" "/win32manifest:app.manifest" "/win32icon:icono.ico" "/r:$sma" /r:System.Core.dll /r:System.Windows.Forms.dll "/resource:payload.bin,payload" Launcher.cs Secrets.g.cs
$ok = ($LASTEXITCODE -eq 0)
Remove-Item (Join-Path $dir 'payload.bin'), (Join-Path $dir 'Secrets.g.cs') -Force -ErrorAction SilentlyContinue
if (-not $ok) { throw 'Fallo la compilacion.' }
Write-Host "LISTO: $out"
