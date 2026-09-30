# Construye dist\ServidorPedidos.exe (un solo archivo, con el servidor cifrado adentro)
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $dir
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
