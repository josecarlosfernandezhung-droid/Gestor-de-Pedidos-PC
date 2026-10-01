<#
====================================================================
 SERVIDOR DE PEDIDOS - Toto Tools
====================================================================
 Este script corre en la PC y hace tres cosas a la vez:

 1) Lee el catalogo de productos desde UN archivo Excel que tu
    eliges (se configura una sola vez desde el Panel, no hay que
    tocar el script) y se lo sirve a CUALQUIER movil que se
    conecte -- ya no hace falta cargar el Excel en cada telefono.
    Cuando modificas ese archivo (precios, productos, cantidades),
    el servidor lo detecta solo y lo actualiza para todos, sin
    reiniciar nada.

 2) Sirve la app "Vendedor" para que los moviles se conecten en:
        http://<IP-de-esta-PC>:8080/vendedor

 3) Sirve el "Panel Receptor" para verlo en esta misma PC en:
        http://localhost:8080/  (se abre solo al iniciar el servidor)

 La lectura del Excel la hace el Panel (en el navegador) usando la
 libreria SheetJS -- debe estar el archivo "xlsx_full_min.js" en
 esta misma carpeta, junto al .ps1. El servidor en PowerShell solo
 vigila el archivo en el disco y guarda el resultado que le manda
 el Panel; por eso el Panel debe quedar abierto (se abre solo).

 CONTROL DE STOCK:
 Si tu Excel tiene una columna de cantidad disponible, el servidor
 descuenta el stock cada vez que se vende algo (sin importar desde
 que telefono) y RECHAZA un pedido si alguien intenta vender mas de
 lo que queda. Si no tienes esa columna, no se controla el stock
 (se permite vender sin limite, como antes).

 IMPORTANTE - el stock vendido NUNCA se pierde al releer el Excel:
 cada vez que el Excel cambia (aunque sea solo el precio o el
 nombre de un producto), el servidor compara la cantidad del Excel
 con la que tenia guardada de la ULTIMA vez que ese numero cambio.
 Si la cantidad del Excel sigue igual, se respetan las ventas
 hechas desde entonces (no "revive" el stock). Solo si TU cambias
 el numero en el Excel (reabasteciste o corregiste el conteo), esa
 cifra se toma como el nuevo punto de partida.

 CONFIGURACION DEL EXCEL (una sola vez, desde el Panel):
 La primera fila del Excel debe ser el encabezado. La primera vez
 que abras el Panel (http://localhost:8080/), te va a pedir la ruta
 completa del archivo (ej: C:\ToTo Tools\catalogo.xlsx) y luego que
 elijas, de una lista con los encabezados reales de tu archivo,
 cual columna es SKU, cual es Nombre, cual es Precio y cual es
 Cantidad/Stock (esta ultima es opcional). Esa configuracion queda
 guardada en "config_excel.json" en esta carpeta, y no te la vuelve
 a preguntar salvo que pulses "Reconfigurar".

 IMPRESION:
 Reutiliza la misma tecnica RAW/winspool del "ayudante de
 impresoras". Cambia $nombreImpresora mas abajo por el nombre EXACTO
 de tu impresora tal como aparece en Windows.

 OTRAS FUNCIONES:
 - Cada vendedor ve, en su propio movil, la lista "Mis pedidos de
   hoy" con boton para reimprimir, cobrar o editar (si sigue
   pendiente), aunque haya perdido la conexion despues de enviarlo.
 - Si al enviar un pedido se cae el WiFi, el movil lo guarda solo y
   lo reintenta cuando vuelva la conexion (no hay que rehacerlo).
 - Desde el Panel de la PC se puede cancelar un pedido pendiente
   (devuelve el stock automaticamente) y cerrar el dia, que archiva
   los pedidos cobrados/cancelados en la carpeta "historial" y dejar
   pedidos.json liviano.

 AJUSTES Y METODO DE PAGO:
   - Panel de la PC -> boton "Ocultar sin stock a vendedores": cuando
     esta ACTIVADO, en el buscador de los vendedores no aparecen los
     productos con cantidad 0 o vacia (se guarda en config_app.json).
   - La PC (caja) no elige el metodo de pago: "Cobrar en Caja" cobra
     con el metodo que dejo el vendedor en el pedido (Efectivo, o
     Transferencia con total x2). El vendedor lo elige en el movil,
     tambien cuando deja el pedido "Pendiente de pago".
   - Recibo estilo TOTOOLS de 32 columnas (ver variables $recibo... al
     inicio del script). En transferencia se imprime el precio por
     transferencia. Tasa USD: casilla del Panel (config_app.json).
   - Movil: Ajustes en un menu aparte, "Mis pedidos de hoy" plegable y
     con lista de productos de cada pedido.
   - Pedido cobrado por un vendedor (desde el movil) = "por revisar":
     sigue visible en el Panel aunque este "solo pendientes", con aviso,
     el monto que debe entregar y el boton "Revisado" (/revisar).
     "Cerrar el dia" no archiva los que faltan por revisar.
   - Mientras la caja no cobre, el vendedor puede editar, cancelar,
     cobrar o cambiar el metodo de pago (se guarda al instante en
     /api/pedidos/N/metodo). Un pedido ya cobrado no se puede anular,
     editar, cambiar ni volver a cobrar (lo valida el servidor).

 NOVEDADES (v14):
   - Etiquetas y codigos de barra: http://localhost:8080/etiquetas (Ajustes).
   - Lista de compras: boton en el aviso de stock bajo y en Ajustes; minimo
     propio por producto (minimos_stock.json).
   - Comprobante de devolucion / cambio / garantia en 32 columnas
     (devoluciones.json).
   - El movil muestra la calidad de la senal con la PC y avisa si cae.
   - Avisos push locales al vendedor: precios, stock, pedidos anulados.
   - Permisos por vendedor desde el Panel (permisos_vendedores.json).

 NOVEDADES (v15):
   - Etiquetas: boton "Nuevos del Excel" marca los productos que aparecieron en el ultimo Excel.
   - Movil (Ajustes): "Descargar lo que va quedando (stock actual)" baja un Excel con las cantidades de ahora.
   - Modo punto de venta en el movil: lo activa el vendedor con la clave de administrador que se pone en la PC
     (menu, "PIN de vendedores"). Ventas (caja), devoluciones, descuentos y reporte de efectivo / transferencia.
   - Mensajes cortos opcionales caja <-> vendedor (nota en el pedido o mensaje suelto). Permiso "mensajes".

 COMO USARLO:
   1. Deja "xlsx_full_min.js", "iniciar.bat" y este script en la
      misma carpeta.
   2. Doble clic en "iniciar.bat" (pide permiso de administrador).
   3. El Panel se abre solo en el navegador. La primera vez,
      configura la ruta del Excel y las columnas (un solo paso).
   4. En los telefonos, abre Chrome y escribe la direccion IP que
      muestra la consola, terminada en /vendedor.
====================================================================
#>

$port = 8080
$nombreImpresora = "CAJA"                                    # <-- nombre EXACTO de tu impresora en Windows
$catalogoPath    = "C:\ToTo Tools\catalogo.xlsx"              # <-- valor SOLO por defecto; se reemplaza al configurar la ruta desde el Panel

# ---- Formato del recibo (ticket) ----
$reciboNombre    = "TOTOOLS"                                 # nombre grande de arriba
$reciboNit       = "84060426826"                             # deja "" para no imprimir el NIT
$reciboPie       = "Los Equipos Electricos tienen una Garantia de 7 dias."   # mensaje final ("" = ninguno)
$reciboAncho     = 32                                        # columnas del papel: 58 mm = 32 | 80 mm = 42 o 48
$reciboConEstilo = $false                                     # letras grandes/negrita (ESC/POS). Si tu impresora imprime simbolos raros, ponlo en $false
$reciboMostrarServicioDescuento = $true                      # imprime SERVICIO y DESCUENTO % igual que el ticket de referencia

# ---- LICENCIA: este script SOLO corre dentro de ServidorPedidos.exe ----
if (-not ("TotoLic.Gate" -as [type]) -or -not $global:TT_APPDIR) {
    Write-Host "Este programa solo puede iniciarse con ServidorPedidos.exe"
    Start-Sleep -Seconds 10
    exit
}
$scriptDir = $global:TT_APPDIR
$pedidosFile = Join-Path $scriptDir "pedidos.json"
$asignadosFile = Join-Path $scriptDir "pedidos_asignados.json"
$alertasFile = Join-Path $scriptDir "alertas_vendedor.json"

# ------------------------------------------------------------------
# Libreria SheetJS (xlsx_full_min.js) -- debe estar en esta carpeta
# ------------------------------------------------------------------
$xlsxLibPath = Join-Path $scriptDir "xlsx_full_min.js"
$global:xlsxLibJs = $global:TT_XLSX
if (-not $global:xlsxLibJs) {
    Write-Host "AVISO: no se encontro 'xlsx_full_min.js' en esta carpeta. El Panel no podra leer el Excel."
    Write-Host "Copia ese archivo junto a este script y reinicia."
}

# ------------------------------------------------------------------
# Icono de la app (PNG, se usa para el manifest / "Agregar a pantalla
# de inicio") y el manifest.json + Service Worker, para que la app
# Vendedor (y el Panel) se puedan agregar como un icono en el
# telefono/escritorio y abrirse de un toque, sin escribir la
# direccion cada vez.
# ------------------------------------------------------------------
$icon192Base64 = @'
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
'@
$icon512Base64 = @'
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
'@
$iconPc192Base64 = @'
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
'@
$iconPc512Base64 = @'
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
'@
$global:iconPc192Bytes = [Convert]::FromBase64String(($iconPc192Base64 -replace '\s',''))
$global:iconPc512Bytes = [Convert]::FromBase64String(($iconPc512Base64 -replace '\s',''))
$global:icon192Bytes = [Convert]::FromBase64String(($icon192Base64 -replace '\s',''))
$global:icon512Bytes = [Convert]::FromBase64String(($icon512Base64 -replace '\s',''))

$global:manifestVendedor = @'
{
  "name": "Toto Tools - Pedidos",
  "short_name": "Pedidos",
  "start_url": "/vendedor",
  "scope": "/",
  "display": "standalone",
  "background_color": "#f1f5f9",
  "theme_color": "#2563eb",
  "icons": [
    { "src": "/icon-192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "/icon-512.png", "sizes": "512x512", "type": "image/png" }
  ]
}
'@

$global:manifestPC = @'
{
  "name": "Toto Tools - Panel de Pedidos",
  "short_name": "Panel Pedidos",
  "start_url": "/",
  "scope": "/",
  "display": "standalone",
  "display_override": ["window-controls-overlay", "standalone"],
  "background_color": "#0f172a",
  "theme_color": "#0f172a",
  "icons": [
    { "src": "/icon-pc-192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "/icon-pc-512.png", "sizes": "512x512", "type": "image/png" }
  ]
}
'@

$global:swJs = @'
// Service Worker: los pedidos y el catalogo SIEMPRE van a la red (nunca se
// sirven de una copia vieja). Lo unico que se cachea son las fotos de
// producto (/foto/...), para que sigan viendose si el telefono pierde la
// conexion con la PC despues de haberlas visto una vez.
const CACHE_FOTOS = 'fotos-catalogo-v1';
// La app del vendedor (y lo que necesita para abrir) se guarda en el telefono:
// si la PC esta apagada, se abre la ultima copia. Solo se usa la copia cuando
// NO hay red; si la PC responde (aunque sea con licencia vencida) manda la PC.
const CACHE_APP = 'app-vendedor-v1';
const ARCHIVOS_APP = ['/vendedor', '/xlsx.js', '/manifest-vendedor.json', '/icon-192.png', '/icon-512.png'];
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(caches.open(CACHE_APP).then((cache) => Promise.all(ARCHIVOS_APP.map(async (u) => {
    try { const r = await fetch(u, { cache: 'no-store' }); if (r && r.ok) await cache.put(u, r); } catch (e) {}
  }))));
});
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
// Al tocar un aviso local se abre (o se trae al frente) la app del vendedor.
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((lista) => {
    for (const c of lista) { if ('focus' in c) return c.focus(); }
    if (self.clients.openWindow) return self.clients.openWindow('/vendedor');
  }));
});
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method === 'GET' && url.pathname.startsWith('/foto/')) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_FOTOS);
      try {
        const resp = await fetch(event.request);
        if (resp && resp.ok) cache.put(event.request, resp.clone());
        return resp;
      } catch (e) {
        const cached = await cache.match(event.request);
        if (cached) return cached;
        throw e;
      }
    })());
    return;
  }
  if (event.request.method === 'GET' && ARCHIVOS_APP.includes(url.pathname) &&
      !(url.pathname === '/vendedor' && url.searchParams.get('cliente') === '1')) {
    event.respondWith((async () => {
      const cache = await caches.open(CACHE_APP);
      try {
        const resp = await fetch(event.request);
        if (resp && resp.ok) cache.put(url.pathname, resp.clone());
        return resp;
      } catch (e) {
        const cached = await cache.match(url.pathname);
        if (cached) return cached;
        throw e;
      }
    })());
    return;
  }
  event.respondWith(fetch(event.request));
});
'@

# ------------------------------------------------------------------
# Estado en memoria: Pedidos
# ------------------------------------------------------------------
$global:pedidos = New-Object System.Collections.ArrayList
$global:nextId = 1

# ---- Vendedores conectados (presencia) y pedidos armados desde la PC ----
# "vendedoresConectados" NO se guarda en disco: es solo presencia (quien
# tiene la app abierta ahora mismo), se reconstruye sola en segundos cuando
# los moviles vuelvan a avisar. "pedidosAsignados" y "alertasVendedor" SI se
# guardan en disco (ver Cargar/Guardar-Asignados y Cargar/Guardar-Alertas)
# para que un pedido "para todos" sin tomar, o una alerta que el vendedor
# aun no vio, sobrevivan si la PC se reinicia o el script se cae.
$global:vendedoresConectados = @{}                          # nombre -> ultima vez visto (DateTime)
$global:segundosVendedorConectado = 20                       # el movil avisa cada 7s; con 20s de margen alcanza
$global:pedidosAsignados = New-Object System.Collections.ArrayList
$global:nextIdAsignado = 1
# Si un pedido "para todos" pasa este tiempo sin que nadie lo tome, se avisa
# en el banner de la PC para que decidas si lo reenvias a alguien puntual.
$global:minutosAvisoSinTomar = 12
# Alertas puntuales para el movil de un vendedor (por ejemplo: el dinero que
# entrego no coincide con lo que el pedido dice que cobro). Se entregan una
# sola vez, como un aviso.
$global:alertasVendedor = New-Object System.Collections.ArrayList
$global:nextIdAlerta = 1

function Cargar-Pedidos {
    if (Test-Path $pedidosFile) {
        try {
            $raw = Get-Content $pedidosFile -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = $raw | ConvertFrom-Json
                foreach ($p in @($data)) {
                    [void]$global:pedidos.Add($p)
                    if ($p.id -ge $global:nextId) { $global:nextId = [int]$p.id + 1 }
                }
            }
            Write-Host "Pedidos anteriores cargados: $($global:pedidos.Count)"
        } catch {
            Write-Host "Aviso: no se pudo leer pedidos.json anterior, se empieza de cero."
        }
    }
}

function Escribir-ArchivoConReintento([string]$ruta, [string]$contenido, [int]$intentos = 5) {
    for ($i = 1; $i -le $intentos; $i++) {
        try {
            Set-Content -Path $ruta -Value $contenido -Encoding UTF8
            return $true
        } catch {
            if ($i -eq $intentos) { Write-Host "Error escribiendo '$ruta' tras $intentos intentos: $_"; return $false }
            Start-Sleep -Milliseconds (80 * $i)   # antivirus/OneDrive sueltan el archivo en milisegundos
        }
    }
    return $false
}

function Guardar-Pedidos {
    Escribir-ArchivoConReintento -ruta $pedidosFile -contenido ($global:pedidos | ConvertTo-Json -Depth 10) | Out-Null
}

function Cargar-Asignados {
    if (Test-Path $asignadosFile) {
        try {
            $raw = Get-Content $asignadosFile -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = $raw | ConvertFrom-Json
                foreach ($a in @($data)) {
                    # "entregadoA" tiene que volver a ser un ArrayList (no un array
                    # fijo de JSON) porque el resto del codigo le hace .Add().
                    $entregadoAList = New-Object System.Collections.ArrayList
                    foreach ($nombre in @($a.entregadoA)) { [void]$entregadoAList.Add($nombre) }
                    $obj = [pscustomobject]@{
                        id            = [int]$a.id
                        vendedor      = $a.vendedor
                        todos         = [bool]$a.todos
                        entregadoA    = $entregadoAList
                        tomadoPor     = $a.tomadoPor
                        horaTomado    = $a.horaTomado
                        items         = $a.items
                        hora          = $a.hora
                        entregado     = [bool]$a.entregado
                        visto         = [bool]$a.visto
                        horaVisto     = $a.horaVisto
                        retirado      = [bool]$a.retirado
                        horaRetirado  = $a.horaRetirado
                        vence         = $a.vence
                    }
                    [void]$global:pedidosAsignados.Add($obj)
                    if ($obj.id -ge $global:nextIdAsignado) { $global:nextIdAsignado = [int]$obj.id + 1 }
                }
            }
            Write-Host "Pedidos asignados anteriores cargados: $($global:pedidosAsignados.Count)"
        } catch {
            Write-Host "Aviso: no se pudo leer pedidos_asignados.json anterior, se empieza de cero."
        }
    }
}

function Guardar-Asignados {
    Escribir-ArchivoConReintento -ruta $asignadosFile -contenido ($global:pedidosAsignados | ConvertTo-Json -Depth 10) | Out-Null
}

function Cargar-Alertas {
    if (Test-Path $alertasFile) {
        try {
            $raw = Get-Content $alertasFile -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = $raw | ConvertFrom-Json
                foreach ($al in @($data)) {
                    [void]$global:alertasVendedor.Add([pscustomobject]@{
                        id       = [int]$al.id
                        vendedor = $al.vendedor
                        mensaje  = $al.mensaje
                        hora     = $al.hora
                        pedidoId = $al.pedidoId
                        tipo     = $al.tipo
                    })
                    if ([int]$al.id -ge $global:nextIdAlerta) { $global:nextIdAlerta = [int]$al.id + 1 }
                }
            }
            Write-Host "Alertas de vendedor anteriores cargadas: $($global:alertasVendedor.Count)"
        } catch {
            Write-Host "Aviso: no se pudo leer alertas_vendedor.json anterior, se empieza de cero."
        }
    }
}

function Guardar-Alertas {
    Escribir-ArchivoConReintento -ruta $alertasFile -contenido ($global:alertasVendedor | ConvertTo-Json -Depth 10) | Out-Null
}

function Guardar-VersionRedimensionada {
    # Guarda UNA version (redimensionada + comprimida) de una foto ya abierta.
    param(
        [System.Drawing.Image]$Img,
        [string]$Destino,
        [int]$AnchoMax,
        [int]$Calidad
    )
    $anchoNuevo = $Img.Width
    $altoNuevo = $Img.Height
    if ($Img.Width -gt $AnchoMax) {
        $escala = $AnchoMax / $Img.Width
        $anchoNuevo = $AnchoMax
        $altoNuevo = [int]([math]::Round($Img.Height * $escala))
    }
    $bmp = New-Object System.Drawing.Bitmap($anchoNuevo, $altoNuevo)
    try {
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.DrawImage($Img, 0, 0, $anchoNuevo, $altoNuevo)
        } finally { $g.Dispose() }

        $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
        $paramsCod = New-Object System.Drawing.Imaging.EncoderParameters(1)
        $paramsCod.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [int64]$Calidad)
        $bmp.Save($Destino, $codec, $paramsCod)
    } finally { $bmp.Dispose() }
}

function Guardar-FotoRedimensionada {
    # Reduce una foto (normalmente a resolucion completa de camara, varios MB)
    # y guarda DOS versiones comprimidas: una chica para la miniatura del
    # buscador y una mas grande para cuando el vendedor la toca para verla
    # ampliada. Si por lo que sea no se puede procesar (foto dañada, formato
    # raro), devuelve $false sin tumbar todo el import.
    param(
        [string]$Origen,
        [string]$DestinoChica,
        [string]$DestinoGrande,
        [int]$AnchoChico = 260,
        [int]$CalidadChica = 65,
        [int]$AnchoGrande = 1080,
        [int]$CalidadGrande = 80
    )
    try {
        $img = [System.Drawing.Image]::FromFile($Origen)
        try {
            Guardar-VersionRedimensionada -Img $img -Destino $DestinoChica -AnchoMax $AnchoChico -Calidad $CalidadChica
            Guardar-VersionRedimensionada -Img $img -Destino $DestinoGrande -AnchoMax $AnchoGrande -Calidad $CalidadGrande
        } finally { $img.Dispose() }
        return $true
    } catch {
        # Si algo fallo al procesarla (foto corrupta, etc.), al menos se
        # copia tal cual (las dos rutas) para no perder la foto por completo.
        try { Copy-Item $Origen $DestinoChica -Force; Copy-Item $Origen $DestinoGrande -Force; return $true } catch { return $false }
    }
}

function Importar-FotosCatalogo {
    # Busca el .zip de "Copia de seguridad" mas reciente que la app del
    # catalogo (Catalogo_de_productos.html) genera en el telefono, y que
    # Jose copio a esta carpeta. Ese zip trae "datos.json" (productos,
    # con su sku) y "fotos/<id>.jpg" (la foto de cada producto, nombrada
    # por su id interno, NO por el sku). Aqui se extraen, se emparejan
    # id->sku usando datos.json, y se guardan en "fotos_catalogo\<sku>.jpg"
    # para poder servirlas por SKU a los vendedores.
    $zips = @(Get-ChildItem -Path $scriptDir -Filter "backup-catalogo-*.zip" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($zips.Count -eq 0) {
        return [pscustomobject]@{ ok = $false; error = "No se encontro ningun archivo 'backup-catalogo-*.zip' en esta carpeta. Desde la app del catalogo, en el telefono, toca 'Copia de seguridad' y copia ese .zip a esta misma carpeta (junto al .ps1)." }
    }
    $zipPath = $zips[0].FullName

    $carpetaFotos = Join-Path $scriptDir "fotos_catalogo"
    if (-not (Test-Path $carpetaFotos)) {
        New-Item -ItemType Directory -Path $carpetaFotos | Out-Null
    }
    $carpetaFotosGrandes = Join-Path $carpetaFotos "grande"
    if (-not (Test-Path $carpetaFotosGrandes)) {
        New-Item -ItemType Directory -Path $carpetaFotosGrandes | Out-Null
    }

    $carpetaTemp = Join-Path $scriptDir "fotos_catalogo_tmp"
    if (Test-Path $carpetaTemp) { Remove-Item $carpetaTemp -Recurse -Force }
    New-Item -ItemType Directory -Path $carpetaTemp | Out-Null

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $carpetaTemp)

        $datosJsonPath = Join-Path $carpetaTemp "datos.json"
        if (-not (Test-Path $datosJsonPath)) {
            throw "El .zip no tiene 'datos.json' (¿es una copia de seguridad valida del catalogo?)."
        }
        $datos = Get-Content $datosJsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $productos = @($datos.products)

        $fotosOrigen = Join-Path $carpetaTemp "fotos"
        Add-Type -AssemblyName System.Drawing
        $copiadas = 0
        $sinFoto = 0
        $conError = 0
        foreach ($p in $productos) {
            $sku = if ($p.sku) { ([string]$p.sku).Trim() } else { "" }
            if (-not $sku) { continue }
            $origen = Join-Path $fotosOrigen ("$($p.id).jpg")
            if (Test-Path $origen) {
                $destinoChica = Join-Path $carpetaFotos ("$sku.jpg")
                $destinoGrande = Join-Path $carpetaFotosGrandes ("$sku.jpg")
                # Las fotos que llegan del telefono suelen venir a resolucion
                # completa de camara (varios MB cada una); con 600+ productos
                # eso es lo que hacia lenta la importacion y pesada de servir
                # a cada movil por la red. Aqui se guardan dos versiones: una
                # chica para el buscador y otra mas grande para cuando el
                # vendedor la toca para verla ampliada.
                if (Guardar-FotoRedimensionada -Origen $origen -DestinoChica $destinoChica -DestinoGrande $destinoGrande) {
                    $copiadas++
                } else {
                    $conError++
                }
            } else {
                $sinFoto++
            }
        }

        Remove-Item $carpetaTemp -Recurse -Force
        return [pscustomobject]@{
            ok = $true
            archivo = $zips[0].Name
            fecha = $zips[0].LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            copiadas = $copiadas
            sinFoto = $sinFoto
            conError = $conError
            totalProductos = $productos.Count
        }
    } catch {
        if (Test-Path $carpetaTemp) { Remove-Item $carpetaTemp -Recurse -Force -ErrorAction SilentlyContinue }
        return [pscustomobject]@{ ok = $false; error = "$_" }
    }
}

function Necesita-Revision($p) {
    # Pedido cobrado por el VENDEDOR (desde el movil) que la caja todavia no reviso.
    return ($p.estado -eq "cobrado" -and $p.cobradoPor -eq "vendedor" -and $p.revisado -eq $false)
}

function Cerrar-Dia {
    # Archiva los pedidos ya cobrados o cancelados en un archivo aparte
    # y deja pedidos.json solo con los que siguen pendientes, para que
    # no crezca para siempre.
    $carpetaHistorial = Join-Path $scriptDir "historial"
    if (-not (Test-Path $carpetaHistorial)) {
        New-Item -ItemType Directory -Path $carpetaHistorial | Out-Null
    }

    # Los cobrados por un vendedor que la caja aun no reviso se quedan en la
    # lista (igual que los pendientes) hasta que se marquen como revisados.
    $paraArchivar = @($global:pedidos | Where-Object { ($_.estado -eq "cobrado" -or $_.estado -eq "cancelado") -and -not (Necesita-Revision $_) })
    $pendientes = @($global:pedidos | Where-Object { ($_.estado -ne "cobrado" -and $_.estado -ne "cancelado") -or (Necesita-Revision $_) })

    if ($paraArchivar.Count -gt 0) {
        $nombreArchivo = "pedidos_" + (Get-Date).ToString("yyyy-MM-dd_HHmmss") + ".json"
        $rutaArchivo = Join-Path $carpetaHistorial $nombreArchivo
        try {
            ($paraArchivar | ConvertTo-Json -Depth 10) | Set-Content -Path $rutaArchivo -Encoding UTF8
        } catch {
            Write-Host "Error archivando historial: $_"
            return $false
        }
    }

    $global:pedidos = New-Object System.Collections.ArrayList
    foreach ($p in $pendientes) { [void]$global:pedidos.Add($p) }
    Guardar-Pedidos
    return $true
}

# ------------------------------------------------------------------
# Metricas: junta los pedidos cobrados de hoy (pedidos.json) con los ya
# archivados (carpeta historial) y calcula: producto mas vendido, hora
# pico y total facturado por vendedor, para el rango de dias pedido.
# ------------------------------------------------------------------
function Obtener-PedidosCobradosDesde([datetime]$desde) {
    $carpetaHistorial = Join-Path $scriptDir "historial"
    $todos = New-Object System.Collections.Generic.List[object]

    if (Test-Path $carpetaHistorial) {
        Get-ChildItem -Path $carpetaHistorial -Filter "pedidos_*.json" -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                $raw = Get-Content $_.FullName -Raw -Encoding UTF8
                if ($raw -and $raw.Trim().Length -gt 0) {
                    $arr = $raw | ConvertFrom-Json
                    foreach ($p in @($arr)) { $todos.Add($p) }
                }
            } catch {}
        }
    }
    foreach ($p in @($global:pedidos)) { $todos.Add($p) }

    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    return @($todos | Where-Object {
        $_.estado -eq "cobrado" -and $_.hora -and
        ([datetime]::ParseExact([string]$_.hora, "yyyy-MM-dd HH:mm:ss", $inv)) -ge $desde
    })
}

function Calcular-Metricas([int]$dias) {
    $desde = (Get-Date).Date.AddDays( - ($dias - 1))
    $pedidos = Obtener-PedidosCobradosDesde $desde
    $inv = [System.Globalization.CultureInfo]::InvariantCulture

    $porProducto = @{}
    $porVendedor = @{}
    $porHora = @{}
    $porMetodo = @{}
    for ($h = 0; $h -le 23; $h++) { $porHora["$h"] = 0 }
    $totalFacturado = 0.0

    foreach ($p in $pedidos) {
        $vend = if ($p.vendedor) { [string]$p.vendedor } else { "Sin nombre" }
        $monto = if ($p.totalCobrado) { [double]$p.totalCobrado } else { [double]$p.totalProductos }
        $totalFacturado += $monto
        if (-not $porVendedor.ContainsKey($vend)) { $porVendedor[$vend] = 0.0 }
        $porVendedor[$vend] += $monto

        $met = if ($p.metodoPago) { [string]$p.metodoPago } else { "Efectivo" }
        if (-not $porMetodo.ContainsKey($met)) { $porMetodo[$met] = 0.0 }
        $porMetodo[$met] += $monto

        try {
            $hora = ([datetime]::ParseExact([string]$p.hora, "yyyy-MM-dd HH:mm:ss", $inv)).Hour
            $porHora["$hora"] += 1
        } catch {}

        foreach ($it in @($p.items)) {
            if (-not $it -or -not $it.nombre) { continue }
            $nombre = [string]$it.nombre
            if (-not $porProducto.ContainsKey($nombre)) { $porProducto[$nombre] = [pscustomobject]@{ cantidad = 0.0; monto = 0.0 } }
            $porProducto[$nombre].cantidad += [double]$it.cantidad
            $porProducto[$nombre].monto += [double]$it.precio * [double]$it.cantidad
        }
    }

    $topProductos = @($porProducto.GetEnumerator() | Sort-Object { $_.Value.cantidad } -Descending | Select-Object -First 10 | ForEach-Object {
        [pscustomobject]@{ nombre = $_.Key; cantidad = $_.Value.cantidad; monto = [math]::Round($_.Value.monto, 2) }
    })
    $porVendedorLista = @($porVendedor.GetEnumerator() | Sort-Object { $_.Value } -Descending | ForEach-Object {
        [pscustomobject]@{ vendedor = $_.Key; monto = [math]::Round($_.Value, 2) }
    })
    $porMetodoLista = @($porMetodo.GetEnumerator() | Sort-Object { $_.Value } -Descending | ForEach-Object {
        [pscustomobject]@{ metodo = $_.Key; monto = [math]::Round($_.Value, 2) }
    })
    $horaPico = ($porHora.GetEnumerator() | Sort-Object { $_.Value } -Descending | Select-Object -First 1)
    $porHoraLista = @(0..23 | ForEach-Object { [pscustomobject]@{ hora = $_; pedidos = $porHora["$_"] } })

    return [pscustomobject]@{
        dias            = $dias
        pedidosCobrados = $pedidos.Count
        totalFacturado  = [math]::Round($totalFacturado, 2)
        topProductos    = $topProductos
        porVendedor     = $porVendedorLista
        porMetodo       = $porMetodoLista
        porHora         = $porHoraLista
        horaPico        = if ($horaPico -and $horaPico.Value -gt 0) { [int]$horaPico.Name } else { $null }
    }
}

Cargar-Pedidos
Cargar-Asignados
Cargar-Alertas

# ------------------------------------------------------------------
# Estado en memoria: Catalogo
# ------------------------------------------------------------------
# La lectura real del .xlsx la hace el navegador (Panel PC) con la
# libreria SheetJS (xlsx_full_min.js, debe estar en esta misma carpeta).
# Este script solo: 1) recuerda la ruta del archivo y el mapeo de
# columnas (config_excel.json), 2) vigila el archivo en disco y avisa
# cuando cambio, 3) sirve los bytes crudos del archivo al panel, y
# 4) recibe de vuelta el catalogo ya parseado (JSON) para guardarlo.
$global:catalogo = New-Object System.Collections.ArrayList
$global:catalogoInfo = [pscustomobject]@{
    archivo          = $catalogoPath
    cargado          = $false
    ultimaCarga      = $null
    error            = $null
    cantidad         = 0
    necesitaRecarga  = $false
    mapeoConfigurado = $false
}

$configExcelPath = Join-Path $scriptDir "config_excel.json"
$global:configExcel = [pscustomobject]@{ ruta = $catalogoPath; mapeo = $null }
$global:catalogoMTimeProcesada = $null

function Cargar-ConfigExcel {
    if (Test-Path $configExcelPath) {
        try {
            $raw = Get-Content $configExcelPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = $raw | ConvertFrom-Json
                if ($data.ruta) { $global:configExcel.ruta = [string]$data.ruta }
                if ($data.mapeo) { $global:configExcel.mapeo = $data.mapeo }
            }
        } catch {
            Write-Host "Aviso: no se pudo leer config_excel.json anterior, se usan valores por defecto."
        }
    }
    $global:catalogoInfo.archivo = $global:configExcel.ruta
    $global:catalogoInfo.mapeoConfigurado = [bool]$global:configExcel.mapeo
}

function Guardar-ConfigExcel {
    try {
        ($global:configExcel | ConvertTo-Json -Depth 10) | Set-Content -Path $configExcelPath -Encoding UTF8
    } catch {
        Write-Host "Error guardando config_excel.json: $_"
    }
}

Cargar-ConfigExcel

# ------------------------------------------------------------------
# Ajustes generales (config_app.json). Por ahora solo uno:
#   ocultarSinStock -> cuando esta activo, la app de los vendedores no
#   muestra en el buscador los productos sin existencias.
# Se cambia desde el Panel de la PC y se recuerda al reiniciar.
# ------------------------------------------------------------------
$configAppPath = Join-Path $scriptDir "config_app.json"
$global:configApp = [pscustomobject]@{ ocultarSinStock = $false; tasaDolar = 0.0; autoservicioDestino = "pc"; wifiSSID = ""; wifiClave = ""; umbralStockBajo = 3; permitirDescuentos = $false }
$global:ipLan = "localhost"

function Cargar-ConfigApp {
    if (Test-Path $configAppPath) {
        try {
            $raw = Get-Content $configAppPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $data = $raw | ConvertFrom-Json
                if ($data.PSObject.Properties.Name -contains 'ocultarSinStock') {
                    $global:configApp.ocultarSinStock = [bool]$data.ocultarSinStock
                }
                if ($data.PSObject.Properties.Name -contains 'tasaDolar') {
                    $global:configApp.tasaDolar = [double]$data.tasaDolar
                }
                if ($data.PSObject.Properties.Name -contains 'umbralStockBajo') {
                    try { $global:configApp.umbralStockBajo = [int]$data.umbralStockBajo } catch { $global:configApp.umbralStockBajo = 3 }
                }
                if ($data.PSObject.Properties.Name -contains 'permitirDescuentos') {
                    $global:configApp.permitirDescuentos = [bool]$data.permitirDescuentos
                }
                if ($data.PSObject.Properties.Name -contains 'autoservicioDestino' -and $data.autoservicioDestino -eq 'vendedor') {
                    $global:configApp.autoservicioDestino = "vendedor"
                }
                if ($data.PSObject.Properties.Name -contains 'wifiSSID') {
                    $global:configApp.wifiSSID = "$($data.wifiSSID)"
                }
                if ($data.PSObject.Properties.Name -contains 'wifiClave') {
                    $global:configApp.wifiClave = "$($data.wifiClave)"
                }
            }
        } catch {
            Write-Host "Aviso: no se pudo leer config_app.json anterior, se usan valores por defecto."
        }
    }
}

function Guardar-ConfigApp {
    Escribir-ArchivoConReintento -ruta $configAppPath -contenido ($global:configApp | ConvertTo-Json -Depth 5) | Out-Null
}

# ------------------------------------------------------------------
# PIN opcional por vendedor. Sin el, cualquiera en la Wi-Fi puede mandar
# pedidos con el nombre de otro vendedor. Con el, el movil debe mandar el
# PIN correcto para crear pedidos o tocar los suyos (metodo/cobrar/cancelar).
# Se guarda por nombre en minusculas: clave -> PIN (texto, ej. "1234").
# ------------------------------------------------------------------
$pinesPath = Join-Path $scriptDir "pines_vendedores.json"
$global:pinesVendedores = @{}

function Cargar-Pines {
    if (Test-Path $pinesPath) {
        try {
            $raw = Get-Content $pinesPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $obj = $raw | ConvertFrom-Json
                $tabla = @{}
                foreach ($prop in $obj.PSObject.Properties) {
                    $v = $prop.Value
                    if ($v -is [string]) {
                        # formato antiguo (clave -> PIN, sin nombre bonito guardado): se
                        # migra usando la clave como nombre, se reescribe al primer guardado.
                        $tabla[$prop.Name] = @{ nombre = $prop.Name; pin = [string]$v }
                    } elseif ($v) {
                        $tabla[$prop.Name] = @{ nombre = [string]$v.nombre; pin = [string]$v.pin }
                    }
                }
                $global:pinesVendedores = $tabla
            }
        } catch { Write-Host "Aviso: no se pudo leer pines_vendedores.json anterior." }
    }
}

function Guardar-Pines {
    Escribir-ArchivoConReintento -ruta $pinesPath -contenido ($global:pinesVendedores | ConvertTo-Json -Depth 3) | Out-Null
}

Cargar-Pines

# Ahora el PIN es obligatorio: un vendedor solo existe si la PC lo creo (ver
# /api/pines y /api/vendedores/crear). Si el nombre no esta registrado, esto
# devuelve $false -- ya no se permite mandar pedidos ni tocar nada a nombre
# de alguien que nadie creo desde la PC.
function Pin-Valido([string]$vendedor, $pinEnviado) {
    if ([string]::IsNullOrWhiteSpace($vendedor)) { return $true }
    $clave = $vendedor.Trim().ToLowerInvariant()
    if (-not $global:pinesVendedores.ContainsKey($clave)) { return $false }
    $pinGuardado = [string]$global:pinesVendedores[$clave].pin
    if ([string]::IsNullOrEmpty($pinGuardado)) { return $true }
    return ([string]$pinEnviado) -eq $pinGuardado
}

Cargar-ConfigApp

function Revisar-CambioCatalogo {
    try {
        $ruta = $global:configExcel.ruta
        if (-not $ruta) {
            $global:catalogoInfo.error = "No se ha configurado la ruta del archivo Excel todavia."
            return
        }
        if (-not (Test-Path $ruta)) {
            $global:catalogoInfo.error = "No se encontro el archivo: $ruta"
            $global:catalogoInfo.necesitaRecarga = $false
            return
        }
        $escritura = (Get-Item $ruta).LastWriteTimeUtc
        if (-not $global:catalogoMTimeProcesada -or $escritura -ne $global:catalogoMTimeProcesada) {
            $global:catalogoInfo.necesitaRecarga = $true
            if (-not $global:catalogoInfo.cargado) { $global:catalogoInfo.error = $null }
        }
    } catch {}
}

Revisar-CambioCatalogo

# ------------------------------------------------------------------
# Impresion RAW (misma tecnica que el "ayudante de impresoras")
# ------------------------------------------------------------------
$codigoRaw = @"
using System;
using System.Runtime.InteropServices;

public class RawPrinterHelper
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
    public class DOCINFOA
    {
        [MarshalAs(UnmanagedType.LPStr)] public string pDocName;
        [MarshalAs(UnmanagedType.LPStr)] public string pOutputFile;
        [MarshalAs(UnmanagedType.LPStr)] public string pDataType;
    }

    [DllImport("winspool.Drv", EntryPoint = "OpenPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool OpenPrinter(string szPrinter, out IntPtr hPrinter, IntPtr pd);

    [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartDocPrinterA", SetLastError = true, CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartDocPrinter(IntPtr hPrinter, int level, [In, MarshalAs(UnmanagedType.LPStruct)] DOCINFOA di);

    [DllImport("winspool.Drv", EntryPoint = "EndDocPrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndDocPrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "StartPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool StartPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "EndPagePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool EndPagePrinter(IntPtr hPrinter);

    [DllImport("winspool.Drv", EntryPoint = "WritePrinter", SetLastError = true, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
    public static extern bool WritePrinter(IntPtr hPrinter, IntPtr pBytes, int dwCount, out int dwWritten);

    public static bool EnviarBytes(string nombreImpresora, byte[] datos)
    {
        IntPtr hPrinter;
        DOCINFOA di = new DOCINFOA();
        di.pDocName = "Pedido Toto Tools";
        di.pDataType = "RAW";
        bool exito = false;

        if (!OpenPrinter(nombreImpresora, out hPrinter, IntPtr.Zero))
            throw new Exception("No se pudo abrir la impresora '" + nombreImpresora + "'. Revisa que el nombre sea exacto.");

        try
        {
            if (StartDocPrinter(hPrinter, 1, di))
            {
                if (StartPagePrinter(hPrinter))
                {
                    IntPtr pUnmanaged = Marshal.AllocCoTaskMem(datos.Length);
                    Marshal.Copy(datos, 0, pUnmanaged, datos.Length);
                    int escritos;
                    exito = WritePrinter(hPrinter, pUnmanaged, datos.Length, out escritos);
                    Marshal.FreeCoTaskMem(pUnmanaged);
                    EndPagePrinter(hPrinter);
                }
                EndDocPrinter(hPrinter);
            }
        }
        finally
        {
            ClosePrinter(hPrinter);
        }
        return exito;
    }
}
"@
if (-not ("RawPrinterHelper" -as [type])) {
    Add-Type -TypeDefinition $codigoRaw -Language CSharp
}

function Quitar-Acentos([string]$s) {
    if (-not $s) { return $s }
    $s = $s.Replace('á','a').Replace('é','e').Replace('í','i').Replace('ó','o').Replace('ú','u')
    $s = $s.Replace('Á','A').Replace('É','E').Replace('Í','I').Replace('Ó','O').Replace('Ú','U')
    $s = $s.Replace('ñ','n').Replace('Ñ','N').Replace('ü','u').Replace('Ü','U')
    return $s
}

function Centrar([string]$texto, [int]$ancho) {
    if ($texto.Length -ge $ancho) { return $texto }
    $espacios = [math]::Floor(($ancho - $texto.Length) / 2)
    return (' ' * $espacios) + $texto
}

# ---- Utilidades de formato para el recibo ----
# Montos como en el ticket de referencia: espacio para los miles y un decimal
# (ej. 4 310.0). Con papel de 32 columnas todo se alinea sin partirse.
function Formato-Monto($n) {
    $v = [math]::Round([double]$n, 1, [System.MidpointRounding]::AwayFromZero)
    return $v.ToString("#,##0.0", [System.Globalization.CultureInfo]::InvariantCulture).Replace(",", " ")
}

function Formato-Cantidad($n) {
    $v = [math]::Round([double]$n, 2, [System.MidpointRounding]::AwayFromZero)
    return $v.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Alinear-Der([string]$texto, [int]$ancho) {
    if ($texto.Length -ge $ancho) { return $texto }
    return (' ' * ($ancho - $texto.Length)) + $texto
}

# "IZQUIERDA ........ DERECHA" ocupando exactamente $ancho columnas
function Linea-LR([string]$izq, [string]$der, [int]$ancho) {
    $espacio = $ancho - $izq.Length - $der.Length
    if ($espacio -lt 1) { $espacio = 1 }
    return $izq + (' ' * $espacio) + $der
}

# Parte un texto largo en varias lineas de maximo $ancho columnas (por palabras)
function Partir-Texto([string]$texto, [int]$ancho) {
    $lineas = @()
    $actual = ""
    foreach ($palabra in ($texto -split '\s+')) {
        if (-not $palabra) { continue }
        if ($palabra.Length -gt $ancho) {
            if ($actual) { $lineas += $actual; $actual = "" }
            while ($palabra.Length -gt $ancho) {
                $lineas += $palabra.Substring(0, $ancho)
                $palabra = $palabra.Substring($ancho)
            }
            $actual = $palabra
        } elseif (-not $actual) {
            $actual = $palabra
        } elseif (($actual.Length + 1 + $palabra.Length) -le $ancho) {
            $actual = $actual + " " + $palabra
        } else {
            $lineas += $actual
            $actual = $palabra
        }
    }
    if ($actual) { $lineas += $actual }
    return $lineas
}

# Comando ESC/POS "modo de impresion": 0 = normal, 24 = negrita + doble alto.
# Si $reciboConEstilo es $false no se manda ningun comando (solo texto).
function Modo-Impresora([int]$n) {
    if (-not $reciboConEstilo) { return "" }
    return ([string][char]27) + "!" + ([string][char]$n)
}

function Generar-TextoRecibo($p) {
    $ancho = [int]$reciboAncho
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder
    $sep = ("- " * [int][math]::Floor($ancho / 2)).TrimEnd()
    $grande = Modo-Impresora 24
    $normal = Modo-Impresora 0

    if ($reciboConEstilo) { [void]$sb.Append(([string][char]27) + "@") }

    # ---- Encabezado ----
    [void]$sb.AppendLine($grande + (Centrar $reciboNombre $ancho) + $normal)
    if ($reciboNit) { [void]$sb.AppendLine($grande + (Centrar ("NIT: " + $reciboNit) $ancho) + $normal) }
    [void]$sb.AppendLine("")

    $fecha = [string]$p.hora
    try { $fecha = ([datetime]::ParseExact([string]$p.hora, "yyyy-MM-dd HH:mm:ss", $inv)).ToString("d/M/yyyy HH:mm", $inv) } catch {}
    [void]$sb.AppendLine("Vendedor: $($p.vendedor)")
    [void]$sb.AppendLine("Folio: $($p.id)")
    [void]$sb.AppendLine("Fecha: $fecha")
    if ([string]$p.estado -ne "cobrado") { [void]$sb.AppendLine("Estado: PENDIENTE DE PAGO") }
    [void]$sb.AppendLine($sep)

    # ---- Productos ----
    # Si se paga por transferencia, el recibo muestra directamente el precio
    # por transferencia de cada producto (nunca el precio base).
    $metodo = [string]$p.metodoPago
    $factor = if ($metodo -eq "Transferencia") { 2 } else { 1 }
    $subtotal = 0.0
    $hayItems = $false
    foreach ($it in @($p.items)) {
        if (-not $it) { continue }
        $hayItems = $true
        $cant = [double]$it.cantidad
        $precioUnit = [double]$it.precio * $factor
        $lineaTotal = $precioUnit * $cant
        $subtotal += $lineaTotal
        foreach ($l in @(Partir-Texto ([string]$it.nombre) $ancho)) { [void]$sb.AppendLine($l) }
        $detalle = "$(Formato-Cantidad $cant) x $(Formato-Monto $precioUnit) $(Formato-Monto $lineaTotal)"
        [void]$sb.AppendLine((Alinear-Der $detalle $ancho))
    }
    if (-not $hayItems) { $subtotal = [double]$p.totalCobrado }
    $total = [math]::Round($subtotal, 2)

    # ---- Totales ----
    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine($grande + (Linea-LR "SUBTOTAL:" (Formato-Monto $total) $ancho) + $normal)
    if ($reciboMostrarServicioDescuento) {
        [void]$sb.AppendLine($grande + (Linea-LR "SERVICIO:" "0.0" $ancho) + $normal)
        [void]$sb.AppendLine($grande + "DESCUENTO %:" + $normal)
    }
    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine($grande + (Linea-LR "TOTAL CUP:" (Formato-Monto $total) $ancho) + $normal)

    $tasa = 0.0
    try { $tasa = [double]$global:configApp.tasaDolar } catch {}
    if ($tasa -gt 0) {
        [void]$sb.AppendLine($grande + (Linea-LR "TOTAL USD:" (Formato-Monto ($total / $tasa)) $ancho) + $normal)
    }
    [void]$sb.AppendLine("")

    $etiquetaPago = if ([string]$p.estado -eq "cobrado") { "PAGADO:" } else { "POR COBRAR:" }
    [void]$sb.AppendLine($grande + (Linea-LR $etiquetaPago (Formato-Monto $total) $ancho) + $normal)
    $textoMetodo = switch ($metodo) {
        "Transferencia" { "Pago por transferencia" }
        "Otro"          { "Otra forma de pago" }
        default         { "Pago en efectivo" }
    }
    [void]$sb.AppendLine((Linea-LR $textoMetodo (Formato-Monto $total) $ancho))

    $recibido = 0.0
    try { $recibido = [double]$p.montoRecibido } catch {}
    if ([string]$p.estado -eq "cobrado" -and $metodo -ne "Transferencia" -and $recibido -gt $total) {
        [void]$sb.AppendLine((Linea-LR "Recibido:" (Formato-Monto $recibido) $ancho))
        [void]$sb.AppendLine((Linea-LR "Cambio:" (Formato-Monto ($recibido - $total)) $ancho))
    }

    # ---- Pie ----
    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine("")
    foreach ($l in @(Partir-Texto $reciboPie $ancho)) { [void]$sb.AppendLine((Centrar $l $ancho)) }
    [void]$sb.Append("`n`n`n`n")
    return (Quitar-Acentos ($sb.ToString()))
}

# ==================================================================
# NUEVAS FUNCIONES: permisos por vendedor, avisos push locales,
# reabastecimiento, ticket de devolucion/garantia, etiquetas.
# ==================================================================
function Enviar-Json($Context, $Obj, [int]$StatusCode = 200, [int]$Depth = 6) {
    Enviar-Respuesta -Context $Context -Body ($Obj | ConvertTo-Json -Depth $Depth) -ContentType "application/json; charset=utf-8" -StatusCode $StatusCode
}

function Leer-CuerpoJson($request) {
    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
    $txt = $reader.ReadToEnd()
    $reader.Close()
    if ($txt -and $txt.Trim().Length -gt 0) { return ($txt | ConvertFrom-Json) }
    return $null
}

# ---------------- Gestor de Almacenes: sincronizacion por WiFi local ----------------
# La app "Gestor de Almacenes" usa ESTE servidor como punto de encuentro (sin internet):
# cada telefono sube sus movimientos (un archivo por dispositivo) y baja los de los
# demas, y aqui se guarda el respaldo por secciones (productos, almacenes, etc.).
# Todo queda en la carpeta "almacen_sync" junto a este script.
$almacenSyncDir = Join-Path $scriptDir "almacen_sync"
$almacenSyncMovDir = Join-Path $almacenSyncDir "mov"
$almacenRespaldoPath = Join-Path $almacenSyncDir "respaldo.json"
$almacenUtf8 = New-Object System.Text.UTF8Encoding($false)

function Almacen-Sha([string]$texto) {
    $md5 = [System.Security.Cryptography.MD5]::Create()
    try { $h = $md5.ComputeHash($almacenUtf8.GetBytes($texto)) } finally { $md5.Dispose() }
    return ([BitConverter]::ToString($h) -replace '-', '').ToLower()
}

function Almacen-Leer([string]$ruta) {
    if (-not (Test-Path $ruta)) { return $null }
    return [System.IO.File]::ReadAllText($ruta, $almacenUtf8)
}

function Almacen-Escribir([string]$ruta, [string]$texto) {
    $dir = Split-Path $ruta -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $ruta + ".tmp"
    [System.IO.File]::WriteAllText($tmp, $texto, $almacenUtf8)
    if (Test-Path $ruta) { Remove-Item $ruta -Force }
    Move-Item $tmp $ruta -Force
}

function Almacen-NombreValido([string]$n) {
    $n = ([string]$n).Trim().ToLower()
    if ($n -match '^[a-z0-9][a-z0-9\-]{0,59}\.json$') { return $n }
    return $null
}

function Leer-CuerpoTexto($request) {
    $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Close() }
}

function Almacen-ListaMovJson {
    $partes = New-Object System.Collections.ArrayList
    if (Test-Path $almacenSyncMovDir) {
        foreach ($f in @(Get-ChildItem -Path $almacenSyncMovDir -Filter "*.json" -File -ErrorAction SilentlyContinue)) {
            $txt = Almacen-Leer $f.FullName
            if ($null -eq $txt) { continue }
            [void]$partes.Add('{"name":"' + $f.Name + '","sha":"' + (Almacen-Sha $txt) + '"}')
        }
    }
    return "[" + ($partes -join ",") + "]"
}

# ---------------- Anti-duplicados PC -> movil ----------------
# Cada envio desde el panel (pedido armado para un vendedor / productos agregados a un pedido)
# lleva una "claveEnvio". Si llega la misma clave otra vez (doble toque, reintento por un corte
# de red) se devuelve la respuesta anterior y NO se crea ni se agrega nada de nuevo.
$global:clavesEnvio = @{}
function Clave-EnvioBuscar([string]$clave) {
    if ([string]::IsNullOrWhiteSpace($clave)) { return $null }
    $limite = (Get-Date).AddMinutes(-20)
    foreach ($k in @($global:clavesEnvio.Keys)) {
        if ($global:clavesEnvio[$k].t -lt $limite) { $global:clavesEnvio.Remove($k) }
    }
    if ($global:clavesEnvio.ContainsKey($clave)) { return [string]$global:clavesEnvio[$clave].resp }
    return $null
}
function Clave-EnvioGuardar([string]$clave, [string]$resp) {
    if ([string]::IsNullOrWhiteSpace($clave)) { return }
    $global:clavesEnvio[$clave] = @{ t = (Get-Date); resp = $resp }
}

function Imprimir-Texto([string]$texto) {
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($texto)
    return [RawPrinterHelper]::EnviarBytes($nombreImpresora, $bytes)
}

function Resumir-Lista($lista, [int]$max = 4) {
    $arr = @($lista)
    $txt = (@($arr | Select-Object -First $max)) -join "; "
    if ($arr.Count -gt $max) { $txt += " y $($arr.Count - $max) mas" }
    return $txt
}

# ---------------- Permisos por vendedor ----------------
# Se guardan en permisos_vendedores.json (clave = nombre en minusculas).
# Un vendedor sin nada guardado tiene TODOS los permisos (como antes).
$permisosPath = Join-Path $scriptDir "permisos_vendedores.json"
$global:permisosVendedores = @{}
$global:catalogoPermisos = @(
    @{ k = 'crearPedidos';  t = 'Crear pedidos';                   d = 'Puede armar y enviar pedidos desde el movil.' },
    @{ k = 'cobrar';        t = 'Cobrar desde el movil';           d = 'Puede marcar pedidos como cobrados.' },
    @{ k = 'cancelar';      t = 'Anular pedidos';                  d = 'Puede cancelar sus pedidos pendientes.' },
    @{ k = 'editar';        t = 'Editar pedidos / metodo de pago'; d = 'Puede modificar un pedido o cambiar el metodo de pago.' },
    @{ k = 'imprimir';      t = 'Imprimir / reimprimir tickets';   d = 'Puede mandar tickets a la impresora de la caja.' },
    @{ k = 'descuentos';    t = 'Cambiar precios (descuentos)';    d = 'Puede bajar el precio de una linea del pedido.' },
    @{ k = 'verStock';      t = 'Ver cantidades de stock';         d = 'Ve cuantas unidades quedan de cada producto.' },
    @{ k = 'misPedidos';    t = 'Entrar a "Mis pedidos de hoy"';   d = 'Ve la lista de sus pedidos del dia.' },
    @{ k = 'ajustes';       t = 'Entrar a Ajustes';                d = 'Abre el menu de Ajustes (catalogo, fotos, PIN).' },
    @{ k = 'modoCliente';   t = 'Usar modo cliente';               d = 'Puede mostrarle el catalogo al cliente.' },
    @{ k = 'asignados';     t = 'Recibir pedidos de la caja';      d = 'Le llegan los pedidos que arma la PC.' },
    @{ k = 'notificaciones'; t = 'Recibir notificaciones';         d = 'Avisos de precios, stock, anulaciones y permisos.' },
    @{ k = 'mensajes';       t = 'Enviar mensajes a la caja';      d = 'Puede escribir una nota en sus pedidos o mandar mensajes cortos a la caja.' }
)

function Permisos-Completos {
    $h = [ordered]@{}
    foreach ($c in $global:catalogoPermisos) { $h[$c.k] = $true }
    return $h
}

function Cargar-Permisos {
    if (Test-Path $permisosPath) {
        try {
            $raw = Get-Content $permisosPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $obj = $raw | ConvertFrom-Json
                $tabla = @{}
                foreach ($prop in $obj.PSObject.Properties) {
                    $h = @{}
                    foreach ($p2 in $prop.Value.PSObject.Properties) { $h[$p2.Name] = [bool]$p2.Value }
                    $tabla[$prop.Name] = $h
                }
                $global:permisosVendedores = $tabla
            }
        } catch { Write-Host "Aviso: no se pudo leer permisos_vendedores.json anterior." }
    }
}

function Guardar-Permisos {
    Escribir-ArchivoConReintento -ruta $permisosPath -contenido ($global:permisosVendedores | ConvertTo-Json -Depth 4) | Out-Null
}

Cargar-Permisos

function Obtener-Permisos([string]$vendedor) {
    $res = Permisos-Completos
    $clave = ([string]$vendedor).Trim().ToLowerInvariant()
    if ($clave -and $global:permisosVendedores.ContainsKey($clave)) {
        $guardado = $global:permisosVendedores[$clave]
        foreach ($c in $global:catalogoPermisos) {
            if ($guardado.ContainsKey($c.k)) { $res[$c.k] = [bool]$guardado[$c.k] }
        }
    }
    return $res
}

function Tiene-Permiso([string]$vendedor, [string]$clave) {
    if ([string]::IsNullOrWhiteSpace($vendedor)) { return $true }
    $p = Obtener-Permisos $vendedor
    return [bool]$p[$clave]
}

# Que permiso exige cada ruta (solo se aplica a peticiones que NO vienen de
# la propia PC y que traen la cabecera X-Vendedor que manda la app movil).
function Permiso-De-Ruta([string]$method, [string]$path) {
    if ($method -eq "POST") {
        if ($path -eq "/api/pedidos") { return "crearPedidos" }
        if ($path -match "^/api/pedidos/\d+/cobrar$") { return "cobrar" }
        if ($path -match "^/api/pedidos/\d+/cancelar$") { return "cancelar" }
        if ($path -match "^/api/pedidos/\d+/metodo$") { return "editar" }
        if ($path -match "^/api/pedidos/\d+/imprimir$") { return "imprimir" }
        if ($path -eq "/api/mensajes") { return "mensajes" }
    } elseif ($method -eq "GET") {
        if ($path -eq "/api/pedidos") { return "misPedidos" }
        if ($path -eq "/api/pedidos/asignados") { return "asignados" }
    }
    return $null
}

# ---------------- Avisos push locales ----------------
# Se apoyan en la cola de alertas que el movil ya consulta cada 7 s
# (/api/alertas). "tipo" permite que el movil muestre el aviso adecuado.
function Agregar-AlertaVendedor([string]$vendedor, [string]$mensaje, $pedidoId = $null, [string]$tipo = "aviso") {
    if ([string]::IsNullOrWhiteSpace($vendedor)) { return }
    $criticos = @('pedido', 'anulado', 'permisos', 'mensaje')
    if (($criticos -notcontains $tipo) -and (-not (Tiene-Permiso $vendedor 'notificaciones'))) { return }
    $al = [pscustomobject]@{
        id       = $global:nextIdAlerta
        vendedor = $vendedor
        mensaje  = $mensaje
        hora     = (Get-Date).ToString("HH:mm:ss")
        pedidoId = $pedidoId
        tipo     = $tipo
    }
    $global:nextIdAlerta++
    [void]$global:alertasVendedor.Add($al)
    $mios = @($global:alertasVendedor | Where-Object { $_.vendedor -eq $vendedor })
    while ($mios.Count -gt 30) {
        [void]$global:alertasVendedor.Remove($mios[0])
        $mios = @($mios | Select-Object -Skip 1)
    }
    Guardar-Alertas
}

function Avisar-A-Todos([string]$mensaje, [string]$tipo = "aviso", [string]$excepto = "") {
    foreach ($k in @($global:pinesVendedores.Keys)) {
        $nom = [string]$global:pinesVendedores[$k].nombre
        if ($excepto -and $nom.ToLowerInvariant() -eq $excepto.Trim().ToLowerInvariant()) { continue }
        Agregar-AlertaVendedor $nom $mensaje $null $tipo
    }
}

# Compara el catalogo anterior con el nuevo (al releer el Excel) y avisa a
# los vendedores de cambios de precio y de disponibilidad.
function Avisar-CambiosCatalogo($anteriorPorSku, $nuevoLista) {
    try {
        if ($anteriorPorSku.Count -eq 0) { return }
        $precios = New-Object System.Collections.ArrayList
        $agotados = New-Object System.Collections.ArrayList
        $repuestos = New-Object System.Collections.ArrayList
        foreach ($n in $nuevoLista) {
            if (-not $n.sku) { continue }
            $a = $anteriorPorSku[[string]$n.sku]
            if (-not $a) { continue }
            if ([math]::Abs([double]$a.precio - [double]$n.precio) -gt 0.001) {
                [void]$precios.Add("$($n.nombre): $(Formato-Monto $a.precio) -> $(Formato-Monto $n.precio)")
            }
            if ($n.stock -ne $null -and $a.stock -ne $null) {
                if ([double]$a.stock -gt 0 -and [double]$n.stock -le 0) {
                    [void]$agotados.Add([string]$n.nombre)
                } elseif ([double]$a.stock -le 0 -and [double]$n.stock -gt 0) {
                    [void]$repuestos.Add("$($n.nombre) ($(Formato-Cantidad $n.stock))")
                }
            }
        }
        if ($precios.Count -gt 0) { Avisar-A-Todos ("Cambio de precio: " + (Resumir-Lista $precios 4)) "precio" }
        if ($agotados.Count -gt 0) { Avisar-A-Todos ("Sin stock: " + (Resumir-Lista $agotados 5)) "stock" }
        if ($repuestos.Count -gt 0) { Avisar-A-Todos ("Vuelve a haber: " + (Resumir-Lista $repuestos 5)) "stock" }
    } catch { Write-Host "Aviso: no se pudieron generar avisos de cambios del catalogo ($_)." }
}

# ---------------- Reabastecimiento (minimos por producto) ----------------
$minimosPath = Join-Path $scriptDir "minimos_stock.json"
$global:minimosStock = @{}

function Cargar-Minimos {
    if (Test-Path $minimosPath) {
        try {
            $raw = Get-Content $minimosPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $obj = $raw | ConvertFrom-Json
                $t = @{}
                foreach ($prop in $obj.PSObject.Properties) { $t[$prop.Name] = [double]$prop.Value }
                $global:minimosStock = $t
            }
        } catch { Write-Host "Aviso: no se pudo leer minimos_stock.json anterior." }
    }
}

function Guardar-Minimos {
    Escribir-ArchivoConReintento -ruta $minimosPath -contenido ($global:minimosStock | ConvertTo-Json -Depth 3) | Out-Null
}

Cargar-Minimos

# Minimo de seguridad de un producto: el suyo propio si se fijo uno, si no el
# umbral general de Ajustes. 0 = sin alerta para ese producto.
function Minimo-Efectivo($p) {
    $sku = [string]$p.sku
    if ($sku -and $global:minimosStock.ContainsKey($sku)) { return [double]$global:minimosStock[$sku] }
    return [double]$global:configApp.umbralStockBajo
}

function Generar-TextoListaCompras($items) {
    $ancho = [int]$reciboAncho
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder
    $sep = ("- " * [int][math]::Floor($ancho / 2)).TrimEnd()
    $grande = Modo-Impresora 24
    $normal = Modo-Impresora 0
    if ($reciboConEstilo) { [void]$sb.Append(([string][char]27) + "@") }
    [void]$sb.AppendLine($grande + (Centrar "LISTA DE COMPRAS" $ancho) + $normal)
    [void]$sb.AppendLine((Centrar $reciboNombre $ancho))
    [void]$sb.AppendLine((Centrar (Get-Date).ToString("d/M/yyyy HH:mm", $inv) $ancho))
    [void]$sb.AppendLine($sep)
    $n = 0
    foreach ($it in @($items)) {
        $cant = Formato-Cantidad $it.cantidad
        foreach ($l in @(Partir-Texto ("[ ] " + $cant + " x " + [string]$it.nombre) $ancho)) { [void]$sb.AppendLine($l) }
        if ($it.stock -ne $null -and "$($it.stock)" -ne "") {
            $det = "    Hay: " + (Formato-Cantidad $it.stock)
            if ($it.minimo -ne $null -and "$($it.minimo)" -ne "") { $det += "  Min: " + (Formato-Cantidad $it.minimo) }
            [void]$sb.AppendLine($det)
        }
        $n++
    }
    [void]$sb.AppendLine($sep)
    [void]$sb.AppendLine("Articulos: $n")
    [void]$sb.Append("`n`n`n`n")
    return (Quitar-Acentos ($sb.ToString()))
}

# ---------------- Etiquetas en la termica (ESC/POS, codigo de barras Code 128 nativo) ----------------
function Escribir-Bytes($ms, [byte[]]$b) { $ms.Write($b, 0, $b.Length) }
function Escribir-Texto($ms, [string]$t) {
    $bytes = [System.Text.Encoding]::ASCII.GetBytes((Quitar-Acentos $t) + "`n")
    $ms.Write($bytes, 0, $bytes.Length)
}

function Generar-BytesEtiquetas($items, $o) {
    $ancho = [int]$reciboAncho
    $maxDots = if ($ancho -le 32) { 384 } else { 576 }
    $ms = New-Object System.IO.MemoryStream
    $feed = 3
    if ($o -and $o.feed -ne $null -and "$($o.feed)" -ne "") { try { $feed = [int]$o.feed } catch { $feed = 3 } }
    if ($feed -lt 0) { $feed = 0 }
    if ($feed -gt 20) { $feed = 20 }
    $corte = [bool]($o -and $o.corte)
    $negocio = if ($o -and $o.negocio) { ([string]$o.negocio).Trim() } else { [string]$reciboNombre }

    Escribir-Bytes $ms ([byte[]]@(27, 64))
    foreach ($it in @($items)) {
        $copias = 1
        try { $copias = [int]$it.copias } catch { $copias = 1 }
        if ($copias -lt 1) { $copias = 1 }
        if ($copias -gt 200) { $copias = 200 }
        $precio = 0.0; try { $precio = [double]$it.precio } catch { $precio = 0.0 }
        $sku = ([string]$it.sku) -replace '[^\x20-\x7E]', ''
        if ($sku.Length -gt 60) { $sku = $sku.Substring(0, 60) }
        for ($c = 0; $c -lt $copias; $c++) {
            Escribir-Bytes $ms ([byte[]]@(27, 97, 1))
            if ($o.mostrarNegocio -and $negocio) {
                Escribir-Bytes $ms ([byte[]]@(27, 69, 1)); Escribir-Texto $ms $negocio; Escribir-Bytes $ms ([byte[]]@(27, 69, 0))
            }
            if ($o.mostrarNombre) {
                foreach ($l in @(Partir-Texto ([string]$it.nombre) $ancho | Select-Object -First 2)) { Escribir-Texto $ms $l }
            }
            if ($o.mostrarPrecio) {
                Escribir-Bytes $ms ([byte[]]@(27, 69, 1, 29, 33, 17))
                Escribir-Texto $ms ("$" + (Formato-Monto $precio))
                Escribir-Bytes $ms ([byte[]]@(29, 33, 0, 27, 69, 0))
            }
            if ($o.mostrarTransf) { Escribir-Texto $ms ("Transferencia: $" + (Formato-Monto ($precio * 2))) }
            if ($o.codigoBarras -and $sku.Length -gt 0) {
                $modulos = 11 * ($sku.Length + 2) + 13
                $wMod = 2
                if ($modulos * 3 -le $maxDots) { $wMod = 3 }
                $hri = if ($o.mostrarSku) { 2 } else { 0 }
                Escribir-Bytes $ms ([byte[]]@(29, 104, 60, 29, 119, $wMod, 29, 72, $hri, 29, 102, 0))
                $datos = [System.Text.Encoding]::ASCII.GetBytes("{B" + $sku)
                Escribir-Bytes $ms ([byte[]]@(29, 107, 73, [byte]$datos.Length))
                Escribir-Bytes $ms $datos
                Escribir-Texto $ms ""
            } elseif ($o.mostrarSku -and $sku.Length -gt 0) {
                Escribir-Texto $ms $sku
            }
            if ($feed -gt 0) { Escribir-Bytes $ms ([byte[]]@(27, 100, [byte]$feed)) }
            if ($corte) { Escribir-Bytes $ms ([byte[]]@(29, 86, 66, 3)) }
        }
    }
    Escribir-Bytes $ms ([byte[]]@(27, 97, 0))
    return $ms.ToArray()
}

# ---------------- Red: IP fija de la PC (para no cambiarla a mano en cada punto de acceso) ----------------
$redConfigPath = Join-Path $scriptDir "red_config.json"
$global:redConfig = [pscustomobject]@{ ultimoOcteto = 0; auto = $false }

function Cargar-RedConfig {
    if (Test-Path $redConfigPath) {
        try {
            $d = (Get-Content $redConfigPath -Raw -Encoding UTF8) | ConvertFrom-Json
            $global:redConfig.ultimoOcteto = [int]$d.ultimoOcteto
            $global:redConfig.auto = [bool]$d.auto
        } catch { Write-Host "Aviso: no se pudo leer red_config.json." }
    }
}
function Guardar-RedConfig {
    Escribir-ArchivoConReintento -ruta $redConfigPath -contenido ($global:redConfig | ConvertTo-Json) | Out-Null
}
Cargar-RedConfig

# Deteccion por WMI y cambio con "netsh": funciona en Windows 7/8/10/11 (no usa cmdlets modernos).
function Obtener-RedActiva {
    try {
        $cfgs = @(Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop | Where-Object { $_.DefaultIPGateway })
        $lista = @()
        foreach ($c in $cfgs) {
            $ad = Get-WmiObject Win32_NetworkAdapter -Filter ("Index=" + $c.Index) -ErrorAction SilentlyContinue
            $nom = if ($ad) { [string]$ad.NetConnectionID } else { "" }
            if (-not $nom) { continue }
            if (($nom + " " + $c.Description) -match 'Virtual|VPN|Loopback|Hyper-V|VMware|VirtualBox|Tailscale|ZeroTier|Npcap|TAP-Windows|Bluetooth') { continue }
            $pos = -1
            $ips = @($c.IPAddress)
            for ($k = 0; $k -lt $ips.Count; $k++) { if ([string]$ips[$k] -match '^\d+\.\d+\.\d+\.\d+$') { $pos = $k; break } }
            if ($pos -lt 0) { continue }
            $mask = [string]@($c.IPSubnet)[$pos]
            $pref = 0
            foreach ($b in $mask.Split('.')) { $n = [int]$b; while ($n -gt 0) { $pref += ($n -band 1); $n = $n -shr 1 } }
            $lista += [pscustomobject]@{
                nombre = $nom; ip = [string]$ips[$pos]; mascara = $mask; prefijo = $pref
                puerta = [string]@($c.DefaultIPGateway)[0]; dhcp = [bool]$c.DHCPEnabled
                dns = @($c.DNSServerSearchOrder | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
            }
        }
        if ($lista.Count -eq 0) { return $null }
        $pref2 = @($lista | Where-Object { $_.nombre -match 'Wi-?Fi|Wireless|WLAN|Inal' })
        if ($pref2.Count -gt 0) { return $pref2[0] }
        return $lista[0]
    } catch { return $null }
}

function Ejecutar-Netsh([string]$argumentos) {
    $salida = (& cmd.exe /c ("netsh " + $argumentos + " 2>&1")) | Out-String
    return @{ ok = ($LASTEXITCODE -eq 0); salida = $salida.Trim() }
}

function Aplicar-IpFija([int]$octeto) {
    $r = Obtener-RedActiva
    if (-not $r) { return @{ ok = $false; error = "No hay una red activa con puerta de enlace. Conectate primero al punto de acceso." } }
    if ($octeto -lt 2 -or $octeto -gt 254) { return @{ ok = $false; error = "El ultimo numero de la IP debe estar entre 2 y 254." } }
    if ($r.prefijo -ne 24) { return @{ ok = $false; error = "Esta red no es 255.255.255.0 (/$($r.prefijo)); por seguridad no se cambia sola. Fijala a mano en Windows." } }
    $p = $r.ip.Split('.')
    $nueva = "$($p[0]).$($p[1]).$($p[2]).$octeto"
    if ($nueva -eq $r.puerta) { return @{ ok = $false; error = "Ese numero es el del router ($($r.puerta)). Elige otro." } }
    if ($nueva -eq $r.ip -and -not $r.dhcp) { $global:ipLan = $nueva; return @{ ok = $true; ip = $nueva; sinCambios = $true } }
    if ($nueva -ne $r.ip) {
        $pingTxt = (& ping.exe -n 1 -w 800 $nueva) | Out-String
        if ($pingTxt -match 'TTL=') { return @{ ok = $false; error = "La IP $nueva ya la usa otro equipo en esta red. Prueba con otro numero." } }
    }
    $nom = $r.nombre
    $res = Ejecutar-Netsh ('interface ip set address name="' + $nom + '" static ' + $nueva + ' 255.255.255.0 ' + $r.puerta)
    if (-not $res.ok) {
        [void](Ejecutar-Netsh ('interface ip set address name="' + $nom + '" source=dhcp'))
        return @{ ok = $false; error = "netsh no pudo fijar la IP: $($res.salida). Abre el servidor como Administrador. Se dejo la red en automatico." }
    }
    $dnsLista = if (@($r.dns).Count -gt 0) { @($r.dns) } else { @($r.puerta) }
    [void](Ejecutar-Netsh ('interface ip set dns name="' + $nom + '" static ' + $dnsLista[0] + ' primary'))
    for ($k = 1; $k -lt [math]::Min($dnsLista.Count, 3); $k++) {
        [void](Ejecutar-Netsh ('interface ip add dns name="' + $nom + '" ' + $dnsLista[$k] + ' index=' + ($k + 1)))
    }
    $global:ipLan = $nueva
    return @{ ok = $true; ip = $nueva }
}

function Volver-IpAutomatica {
    $r = Obtener-RedActiva
    if (-not $r) { return @{ ok = $false; error = "No hay una red activa." } }
    $res = Ejecutar-Netsh ('interface ip set address name="' + $r.nombre + '" source=dhcp')
    [void](Ejecutar-Netsh ('interface ip set dns name="' + $r.nombre + '" source=dhcp'))
    if (-not $res.ok) { return @{ ok = $false; error = "netsh no pudo: $($res.salida). Abre el servidor como Administrador." } }
    [void](& ipconfig.exe /renew)
    return @{ ok = $true }
}

# ---------------- Devoluciones / garantias ----------------
$devolucionesFile = Join-Path $scriptDir "devoluciones.json"
$global:devoluciones = New-Object System.Collections.ArrayList
$global:nextIdDevolucion = 1

function Cargar-Devoluciones {
    if (Test-Path $devolucionesFile) {
        try {
            $raw = Get-Content $devolucionesFile -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                foreach ($d in @($raw | ConvertFrom-Json)) {
                    [void]$global:devoluciones.Add($d)
                    if ([int]$d.id -ge $global:nextIdDevolucion) { $global:nextIdDevolucion = [int]$d.id + 1 }
                }
            }
        } catch { Write-Host "Aviso: no se pudo leer devoluciones.json anterior." }
    }
}

function Guardar-Devoluciones {
    while ($global:devoluciones.Count -gt 500) { $global:devoluciones.RemoveAt(0) }
    Escribir-ArchivoConReintento -ruta $devolucionesFile -contenido ($global:devoluciones | ConvertTo-Json -Depth 8) | Out-Null
}

Cargar-Devoluciones

function Generar-TextoDevolucion($d) {
    $ancho = [int]$reciboAncho
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $sb = New-Object System.Text.StringBuilder
    $sep = ("- " * [int][math]::Floor($ancho / 2)).TrimEnd()
    $grande = Modo-Impresora 24
    $normal = Modo-Impresora 0
    if ($reciboConEstilo) { [void]$sb.Append(([string][char]27) + "@") }

    [void]$sb.AppendLine($grande + (Centrar $reciboNombre $ancho) + $normal)
    if ($reciboNit) { [void]$sb.AppendLine((Centrar ("NIT: " + $reciboNit) $ancho)) }
    [void]$sb.AppendLine("")
    $titulo = switch ([string]$d.tipo) { "garantia" { "GARANTIA" } "cambio" { "CAMBIO" } default { "DEVOLUCION" } }
    [void]$sb.AppendLine($grande + (Centrar ("COMPROBANTE DE " + $titulo) $ancho) + $normal)
    [void]$sb.AppendLine("")

    $fecha = [string]$d.hora
    try { $fecha = ([datetime]::ParseExact([string]$d.hora, "yyyy-MM-dd HH:mm:ss", $inv)).ToString("d/M/yyyy HH:mm", $inv) } catch {}
    [void]$sb.AppendLine("Folio: D-$($d.id)")
    [void]$sb.AppendLine("Fecha: $fecha")
    if ($d.pedidoId) { [void]$sb.AppendLine("Pedido original: #$($d.pedidoId)") }
    if ($d.cliente) { foreach ($l in @(Partir-Texto ("Cliente: " + [string]$d.cliente) $ancho)) { [void]$sb.AppendLine($l) } }
    if ($d.atendio) { [void]$sb.AppendLine("Atendio: $($d.atendio)") }
    [void]$sb.AppendLine($sep)

    [void]$sb.AppendLine("PRODUCTOS RECIBIDOS:")
    foreach ($it in @($d.items)) {
        $cant = Formato-Cantidad $it.cantidad
        foreach ($l in @(Partir-Texto ($cant + " x " + [string]$it.nombre) $ancho)) { [void]$sb.AppendLine($l) }
        if ([double]$it.precio -gt 0) {
            $sub = [double]$it.precio * [double]$it.cantidad
            [void]$sb.AppendLine((Alinear-Der ("$" + (Formato-Monto $sub)) $ancho))
        }
    }
    [void]$sb.AppendLine($sep)

    if ([double]$d.total -gt 0) { [void]$sb.AppendLine($grande + (Linea-LR "TOTAL:" ("$" + (Formato-Monto $d.total)) $ancho) + $normal) }
    if ($d.metodoReembolso) { [void]$sb.AppendLine("Reembolso: $($d.metodoReembolso)") }
    if ($d.entregado) {
        [void]$sb.AppendLine("Entregado en cambio:")
        foreach ($l in @(Partir-Texto ([string]$d.entregado) $ancho)) { [void]$sb.AppendLine($l) }
    }
    if ($d.diferencia -ne $null -and "$($d.diferencia)" -ne "" -and [double]$d.diferencia -ne 0) {
        $dif = [double]$d.diferencia
        $txtDif = if ($dif -gt 0) { "Diferencia a cobrar:" } else { "Diferencia a devolver:" }
        [void]$sb.AppendLine((Linea-LR $txtDif ("$" + (Formato-Monto ([math]::Abs($dif)))) $ancho))
    }
    if ($d.motivo) {
        [void]$sb.AppendLine("Motivo:")
        foreach ($l in @(Partir-Texto ([string]$d.motivo) $ancho)) { [void]$sb.AppendLine($l) }
    }
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Cliente:")
    [void]$sb.AppendLine(("_" * ($ancho - 2)))
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Autoriza:")
    [void]$sb.AppendLine(("_" * ($ancho - 2)))
    [void]$sb.AppendLine("")
    if ($reciboPie) { foreach ($l in @(Partir-Texto $reciboPie $ancho)) { [void]$sb.AppendLine((Centrar $l $ancho)) } }
    [void]$sb.Append("`n`n`n`n")
    return (Quitar-Acentos ($sb.ToString()))
}

# ------------------------------------------------------------------
# Productos NUEVOS del ultimo Excel (para imprimir sus etiquetas)
# ------------------------------------------------------------------
# Cada vez que se carga el catalogo se comparan sus productos con los que ya
# se habian visto antes (skus_conocidos.json). Los que no estaban se guardan
# en productos_nuevos.json y la pagina de etiquetas los marca con un toque
# ("Nuevos del Excel"). Si un Excel no trae nada nuevo, la lista anterior se
# conserva (asi no se pierde por volver a cargar el mismo archivo).
# La primera vez solo se "aprende" el catalogo: no se marca nada como nuevo.
$skusConocidosPath = Join-Path $scriptDir "skus_conocidos.json"
$productosNuevosPath = Join-Path $scriptDir "productos_nuevos.json"
$global:skusConocidos = @{}
$global:skusConocidosListos = $false
$global:productosNuevos = [pscustomobject]@{ fecha = ""; skus = @() }

function Clave-Producto-Nuevo($p) {
    if (-not [string]::IsNullOrWhiteSpace([string]$p.sku)) { return [string]$p.sku }
    return ("n:" + [string]$p.nombre)
}

function Cargar-ProductosNuevos {
    try {
        if (Test-Path $skusConocidosPath) {
            $raw = Get-Content $skusConocidosPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                foreach ($s in @($raw | ConvertFrom-Json)) { if ($s) { $global:skusConocidos[[string]$s] = $true } }
                $global:skusConocidosListos = ($global:skusConocidos.Count -gt 0)
            }
        }
    } catch { Write-Host "Aviso: no se pudo leer skus_conocidos.json anterior." }
    try {
        if (Test-Path $productosNuevosPath) {
            $raw2 = Get-Content $productosNuevosPath -Raw -Encoding UTF8
            if ($raw2 -and $raw2.Trim().Length -gt 0) {
                $o = $raw2 | ConvertFrom-Json
                $listaN = New-Object System.Collections.ArrayList
                foreach ($s2 in @($o.skus)) { if ($s2) { [void]$listaN.Add([string]$s2) } }
                $global:productosNuevos = [pscustomobject]@{ fecha = [string]$o.fecha; skus = @($listaN) }
            }
        }
    } catch { Write-Host "Aviso: no se pudo leer productos_nuevos.json anterior." }
}

Cargar-ProductosNuevos

function Registrar-ProductosNuevos($nuevoLista) {
    try {
        if (@($nuevoLista).Count -eq 0) { return }   # un catalogo vacio (archivo a medio guardar) no cuenta
        $nuevosAhora = New-Object System.Collections.ArrayList
        foreach ($n in $nuevoLista) {
            $kn = Clave-Producto-Nuevo $n
            if (-not $global:skusConocidos.ContainsKey($kn)) {
                [void]$nuevosAhora.Add($kn)
                $global:skusConocidos[$kn] = $true
            }
        }
        $primeraVez = (-not $global:skusConocidosListos)
        if ((-not $primeraVez) -and $nuevosAhora.Count -gt 0) {
            $global:productosNuevos = [pscustomobject]@{ fecha = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); skus = @($nuevosAhora) }
            Escribir-ArchivoConReintento -ruta $productosNuevosPath -contenido (ConvertTo-Json -InputObject $global:productosNuevos -Depth 4) | Out-Null
            Write-Host "Productos nuevos detectados en el Excel: $($nuevosAhora.Count)"
        }
        $global:skusConocidosListos = $true
        if ($primeraVez -or $nuevosAhora.Count -gt 0) {
            Escribir-ArchivoConReintento -ruta $skusConocidosPath -contenido (ConvertTo-Json -InputObject @($global:skusConocidos.Keys) -Depth 2 -Compress) | Out-Null
        }
    } catch { Write-Host "Aviso: no se pudieron registrar los productos nuevos ($_)." }
}

# ------------------------------------------------------------------
# MODO PUNTO DE VENTA del movil
# ------------------------------------------------------------------
# El administrador pone una "clave de administrador" desde la PC (menu, PIN de
# vendedores). Con esa clave, en Ajustes del movil de un vendedor se activa el
# modo punto de venta: historial de ventas, devoluciones, descuentos y reporte
# de efectivo / transferencia. Todo sigue pasando por este servidor, asi que
# la PC ve las mismas ventas y el mismo stock en tiempo real.
# La clave se guarda solo como hash (pos_clave.json). Cada movil activado recibe
# un token propio (tambien guardado como hash); si se cambia o se quita la
# clave, todos los tokens se borran y los moviles pierden el modo.
$posClavePath = Join-Path $scriptDir "pos_clave.json"
$global:posClaveHash = ""
$global:posTokens = @{}
$global:posIntentos = @{}

function Pos-Hash([string]$texto) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("toto-pos|" + $texto)) } finally { $sha.Dispose() }
    return ([BitConverter]::ToString($h) -replace '-', '').ToLower()
}

function Cargar-Pos {
    if (Test-Path $posClavePath) {
        try {
            $raw = Get-Content $posClavePath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $o = $raw | ConvertFrom-Json
                if ($o.hash) { $global:posClaveHash = [string]$o.hash }
                foreach ($t in @($o.tokens)) {
                    if ($t -and $t.h) { $global:posTokens[[string]$t.h] = [string]$t.v }
                }
            }
        } catch { Write-Host "Aviso: no se pudo leer pos_clave.json anterior." }
    }
}

function Guardar-Pos {
    $lista = New-Object System.Collections.ArrayList
    foreach ($k in @($global:posTokens.Keys)) { [void]$lista.Add([pscustomobject]@{ h = [string]$k; v = [string]$global:posTokens[$k] }) }
    $obj = [pscustomobject]@{ hash = [string]$global:posClaveHash; tokens = @($lista) }
    Escribir-ArchivoConReintento -ruta $posClavePath -contenido (ConvertTo-Json -InputObject $obj -Depth 4) | Out-Null
}

Cargar-Pos

function Pos-VendedorDe($request) {
    $v = ""
    try { $v = [uri]::UnescapeDataString([string]$request.Headers["X-Vendedor"]) } catch { $v = "" }
    return $v.Trim()
}

# true si la peticion trae un token de punto de venta valido para ESE vendedor.
function Pos-Token-Valido($request, [string]$vendedor) {
    $tok = ""
    try { $tok = [string]$request.Headers["X-Pos-Token"] } catch { $tok = "" }
    if ([string]::IsNullOrWhiteSpace($tok) -or [string]::IsNullOrWhiteSpace($vendedor)) { return $false }
    $h = Pos-Hash $tok
    if (-not $global:posTokens.ContainsKey($h)) { return $false }
    return ([string]$global:posTokens[$h]) -eq $vendedor.Trim().ToLowerInvariant()
}

# Despues de 5 claves erroneas seguidas, esa IP espera 60 segundos.
function Pos-Bloqueado([string]$ip) {
    if ($global:posIntentos.ContainsKey($ip)) {
        $e = $global:posIntentos[$ip]
        if ($e.hasta -and ((Get-Date) -lt $e.hasta)) { return [int][math]::Ceiling(($e.hasta - (Get-Date)).TotalSeconds) }
    }
    return 0
}

function Pos-Fallo([string]$ip) {
    if (-not $global:posIntentos.ContainsKey($ip)) { $global:posIntentos[$ip] = @{ n = 0; hasta = $null } }
    $e = $global:posIntentos[$ip]
    $e.n = [int]$e.n + 1
    if ($e.n -ge 5) { $e.hasta = (Get-Date).AddSeconds(60); $e.n = 0 }
}

function Clave-Linea($it) {
    if (-not [string]::IsNullOrWhiteSpace([string]$it.sku)) { return [string]$it.sku }
    return [string]$it.nombre
}

# Cuanto se ha devuelto ya de cada venta: clave "idPedido|horaPedido" -> (producto -> cantidad).
function Indice-Devoluciones {
    $idx = @{}
    foreach ($x in @($global:devoluciones)) {
        if (-not $x.pedidoId) { continue }
        $hx = if ($x.pedidoHora) { [string]$x.pedidoHora } else { "" }
        $kp = ([string]$x.pedidoId) + "|" + $hx
        if (-not $idx.ContainsKey($kp)) { $idx[$kp] = @{} }
        foreach ($xi in @($x.items)) {
            $kl = Clave-Linea $xi
            if (-not $idx[$kp].ContainsKey($kl)) { $idx[$kp][$kl] = 0.0 }
            $idx[$kp][$kl] = [double]$idx[$kp][$kl] + [double]$xi.cantidad
        }
    }
    return $idx
}

function Mapa-Devuelto($idx, [string]$idPed, [string]$horaPed) {
    $res = @{}
    $claves = @(($idPed + "|" + $horaPed))
    if ($horaPed) { $claves += ($idPed + "|") }   # devoluciones antiguas (hechas desde la PC) sin hora de pedido
    foreach ($kp in $claves) {
        if ($idx.ContainsKey($kp)) {
            foreach ($kl in @($idx[$kp].Keys)) {
                if (-not $res.ContainsKey($kl)) { $res[$kl] = 0.0 }
                $res[$kl] = [double]$res[$kl] + [double]$idx[$kp][$kl]
            }
        }
    }
    return $res
}

# Los archivos de la carpeta "historial" no se reescriben despues de cerrar el dia: se leen una vez y se
# recuerdan en memoria (clave = ruta + tamano), asi consultar el historial cada pocos segundos no castiga el disco.
$global:histCache = @{}
function Leer-HistorialCache($archivo) {
    $clave = $archivo.FullName + "|" + $archivo.Length
    if ($global:histCache.ContainsKey($clave)) { return $global:histCache[$clave] }
    $lista = @()
    try {
        $raw = Get-Content $archivo.FullName -Raw -Encoding UTF8
        if ($raw -and $raw.Trim().Length -gt 0) { $lista = @($raw | ConvertFrom-Json) }
    } catch { $lista = @() }
    if ($global:histCache.Count -gt 120) { $global:histCache = @{} }
    $global:histCache[$clave] = $lista
    return $lista
}

# Pedidos cobrados desde una fecha. Solo abre los archivos del historial que
# pueden tener ventas de ese rango (el nombre del archivo lleva la fecha de cierre).
function Pos-PedidosCobradosDesde([datetime]$desde) {
    $carpeta = Join-Path $scriptDir "historial"
    $todos = New-Object System.Collections.Generic.List[object]
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    if (Test-Path $carpeta) {
        foreach ($f in @(Get-ChildItem -Path $carpeta -Filter "pedidos_*.json" -ErrorAction SilentlyContinue)) {
            if ($f.Name -match '^pedidos_(\d{4}-\d{2}-\d{2})_') {
                $fd = [datetime]::MinValue
                if ([datetime]::TryParseExact($Matches[1], "yyyy-MM-dd", $inv, [System.Globalization.DateTimeStyles]::None, [ref]$fd)) {
                    if ($fd -lt $desde.Date) { continue }
                }
            }
            foreach ($p in @(Leer-HistorialCache $f)) { if ($p) { $todos.Add($p) } }
        }
    }
    foreach ($p in @($global:pedidos)) { $todos.Add($p) }
    return @($todos | Where-Object {
        $ok = $false
        try { $ok = ($_.estado -eq "cobrado") -and $_.hora -and (([datetime]::ParseExact([string]$_.hora, "yyyy-MM-dd HH:mm:ss", $inv)) -ge $desde) } catch { $ok = $false }
        $ok
    })
}

# Busca una venta por folio y hora, tanto entre las de hoy como entre las archivadas.
function Buscar-PedidoCualquiera([int]$id, [string]$hora) {
    foreach ($p in @($global:pedidos)) {
        if (([int]$p.id -eq $id) -and ((-not $hora) -or ([string]$p.hora -eq $hora))) { return $p }
    }
    $carpeta = Join-Path $scriptDir "historial"
    if (Test-Path $carpeta) {
        foreach ($f in @(Get-ChildItem -Path $carpeta -Filter "pedidos_*.json" -ErrorAction SilentlyContinue)) {
            foreach ($x in @(Leer-HistorialCache $f)) {
                if ($x -and ([int]$x.id -eq $id) -and ((-not $hora) -or ([string]$x.hora -eq $hora))) { return $x }
            }
        }
    }
    return $null
}

# Historial de ventas cobradas + reporte de caja (efectivo / transferencia) del rango.
# $quien = "" (todos los vendedores) o el nombre del vendedor en minusculas.
# La parte en transferencia ya viene con el x2 (es el dinero que de verdad entro).
function Calcular-PosVentas([int]$dias, [string]$quien) {
    $desde = (Get-Date).Date.AddDays( - ($dias - 1))
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $todos = @(Pos-PedidosCobradosDesde $desde)
    if ($quien) { $todos = @($todos | Where-Object { ([string]$_.vendedor).Trim().ToLowerInvariant() -eq $quien }) }
    $todos = @($todos | Sort-Object { [string]$_.hora } -Descending)
    $idxDev = Indice-Devoluciones

    $devs = New-Object System.Collections.ArrayList
    foreach ($x in @($global:devoluciones)) {
        $okFecha = $false
        try { $okFecha = (([datetime]::ParseExact([string]$x.hora, "yyyy-MM-dd HH:mm:ss", $inv)) -ge $desde) } catch { $okFecha = $false }
        if (-not $okFecha) { continue }
        if ($quien -and (([string]$x.atendio).Trim().ToLowerInvariant() -ne $quien)) { continue }
        [void]$devs.Add($x)
    }

    $efIn = 0.0; $trIn = 0.0; $otIn = 0.0; $baseTot = 0.0; $descTot = 0.0
    $porVend = @{}
    $lista = New-Object System.Collections.ArrayList
    foreach ($p in $todos) {
        $m = if ($p.metodoPago) { [string]$p.metodoPago } else { "Efectivo" }
        $tp = 0.0; try { $tp = [double]$p.totalProductos } catch { $tp = 0.0 }
        $tc = $tp
        try { if ($p.totalCobrado) { $tc = [double]$p.totalCobrado } } catch {}
        $e = 0.0; $t = 0.0; $o = 0.0
        if ($m -eq "Transferencia") { $t = $tc }
        elseif ($m -eq "Combinado") {
            $pe = 0.0; try { $pe = [double]$p.pagoEfectivo } catch {}
            $pt = 0.0; try { $pt = [double]$p.pagoTransferencia } catch {}
            $e = $pe; $t = $pt * 2
        }
        elseif ($m -eq "Otro") { $o = $tc }
        else { $e = $tc }
        $efIn += $e; $trIn += $t; $otIn += $o; $baseTot += $tp

        $descP = 0.0
        foreach ($it in @($p.items)) {
            try {
                if (($it.PSObject.Properties.Name -contains 'precioLista') -and ($null -ne $it.precioLista)) {
                    $dl = ([double]$it.precioLista - [double]$it.precio) * [double]$it.cantidad
                    if ($dl -gt 0.0001) { $descP += $dl }
                }
            } catch {}
        }
        $descTot += $descP

        $vn = if ($p.vendedor) { [string]$p.vendedor } else { "Sin nombre" }
        $kv = $vn.Trim().ToLowerInvariant()
        if (-not $porVend.ContainsKey($kv)) { $porVend[$kv] = [pscustomobject]@{ vendedor = $vn; ventas = 0; efectivo = 0.0; transferencia = 0.0; otro = 0.0 } }
        $pv = $porVend[$kv]
        $pv.ventas = [int]$pv.ventas + 1
        $pv.efectivo = [double]$pv.efectivo + $e
        $pv.transferencia = [double]$pv.transferencia + $t
        $pv.otro = [double]$pv.otro + $o

        if ($lista.Count -lt 200) {
            $itemsSlim = New-Object System.Collections.ArrayList
            foreach ($it2 in @($p.items)) {
                [void]$itemsSlim.Add([pscustomobject]@{ sku = [string]$it2.sku; nombre = [string]$it2.nombre; cantidad = [double]$it2.cantidad; precio = [double]$it2.precio })
            }
            [void]$lista.Add([pscustomobject]@{
                id = [int]$p.id
                hora = [string]$p.hora
                vendedor = $vn
                metodoPago = $m
                totalProductos = [math]::Round($tp, 2)
                totalCobrado = [math]::Round($tc, 2)
                pagoEfectivo = $p.pagoEfectivo
                pagoTransferencia = $p.pagoTransferencia
                descuento = [math]::Round($descP, 2)
                items = @($itemsSlim)
                devuelto = (Mapa-Devuelto $idxDev ([string]$p.id) ([string]$p.hora))
            })
        }
    }

    $devEf = 0.0; $devTr = 0.0
    foreach ($x in $devs) {
        if ($x.PSObject.Properties.Name -contains 'reembolsoEfectivo') {
            $devEf += [double]$x.reembolsoEfectivo
            $devTr += [double]$x.reembolsoTransferencia
        } elseif ([string]$x.tipo -eq "devolucion") {
            # Devolucion hecha desde la PC (antes de este modo): se toma el total y, si dice transferencia, x2.
            $tt = 0.0; try { $tt = [double]$x.total } catch { $tt = 0.0 }
            if (([string]$x.metodoReembolso).ToLower() -match "transf") { $devTr += $tt * 2 } else { $devEf += $tt }
        }
    }

    $vendLista = @($porVend.Values | Sort-Object { [double]$_.efectivo + [double]$_.transferencia + [double]$_.otro } -Descending | ForEach-Object {
        [pscustomobject]@{ vendedor = $_.vendedor; ventas = $_.ventas; efectivo = [math]::Round($_.efectivo, 2); transferencia = [math]::Round($_.transferencia, 2); otro = [math]::Round($_.otro, 2) }
    })

    $netoEf = $efIn - $devEf
    $netoTr = $trIn - $devTr
    return @{
        ok = $true
        reporte = [pscustomobject]@{
            dias = $dias
            ventas = $todos.Count
            totalProductos = [math]::Round($baseTot, 2)
            descuentos = [math]::Round($descTot, 2)
            efectivoEntrada = [math]::Round($efIn, 2)
            transferenciaEntrada = [math]::Round($trIn, 2)
            otroEntrada = [math]::Round($otIn, 2)
            devoluciones = $devs.Count
            devolucionEfectivo = [math]::Round($devEf, 2)
            devolucionTransferencia = [math]::Round($devTr, 2)
            netoEfectivo = [math]::Round($netoEf, 2)
            netoTransferencia = [math]::Round($netoTr, 2)
            netoTotal = [math]::Round($netoEf + $netoTr + $otIn, 2)
            porVendedor = $vendLista
        }
        pedidos = @($lista)
        hayMas = ($todos.Count -gt 200)
    }
}

# Devolucion hecha desde el movil en modo punto de venta. Comprueba contra la
# venta original que no se devuelva mas de lo vendido, reintegra el stock (si se
# pide), guarda el registro en devoluciones.json (el mismo de la PC) e imprime el
# comprobante en la impresora de la caja. Devuelve @{ codigo; cuerpo }.
function Registrar-DevolucionPos($d, [string]$vendedor) {
    $idPed = 0; try { $idPed = [int]$d.pedidoId } catch { $idPed = 0 }
    $horaPed = [string]$d.pedidoHora
    $pedido = Buscar-PedidoCualquiera $idPed $horaPed
    if ((-not $pedido) -or ([string]$pedido.estado -ne "cobrado")) {
        return @{ codigo = 404; cuerpo = @{ ok = $false; error = "No se encontro esa venta cobrada." } }
    }
    $idxDev = Indice-Devoluciones
    $yaDev = Mapa-Devuelto $idxDev ([string]$idPed) ([string]$pedido.hora)

    $met = if ($pedido.metodoPago) { [string]$pedido.metodoPago } else { "Efectivo" }
    $factorItems = if ($met -eq "Transferencia") { 2.0 } else { 1.0 }   # el cliente pago el precio x2 por unidad

    $itemsDev = New-Object System.Collections.ArrayList
    foreach ($it in @($d.items)) {
        $cant = 0.0; try { $cant = [double]$it.cantidad } catch { $cant = 0.0 }
        if ($cant -le 0) { continue }
        $kl = Clave-Linea $it
        $linea = $null
        foreach ($cand in @($pedido.items)) { if ((Clave-Linea $cand) -eq $kl) { $linea = $cand; break } }
        if (-not $linea) {
            return @{ codigo = 400; cuerpo = @{ ok = $false; error = "Ese producto no esta en la venta." } }
        }
        $ya = if ($yaDev.ContainsKey($kl)) { [double]$yaDev[$kl] } else { 0.0 }
        $yaPedido = 0.0
        foreach ($prev in $itemsDev) { if ((Clave-Linea $prev) -eq $kl) { $yaPedido += [double]$prev.cantidad } }
        $resta = [double]$linea.cantidad - $ya
        if (($cant + $yaPedido) -gt ($resta + 0.0001)) {
            return @{ codigo = 409; cuerpo = @{ ok = $false; error = ("De " + [string]$linea.nombre + " solo se pueden devolver " + (Formato-Cantidad $resta) + ".") } }
        }
        [void]$itemsDev.Add([pscustomobject]@{ sku = [string]$linea.sku; nombre = [string]$linea.nombre; cantidad = $cant; precio = [math]::Round(([double]$linea.precio * $factorItems), 2) })
    }
    if ($itemsDev.Count -eq 0) {
        return @{ codigo = 400; cuerpo = @{ ok = $false; error = "Elige al menos un producto a devolver." } }
    }

    $totalDev = 0.0
    foreach ($i in $itemsDev) { $totalDev += [double]$i.precio * [double]$i.cantidad }
    $totalDev = [math]::Round($totalDev, 2)

    # Cuanto se le devuelve al cliente y por donde (segun como pago la venta).
    $baseDev = 0.0
    foreach ($i in $itemsDev) { $baseDev += ([double]$i.precio / $factorItems) * [double]$i.cantidad }
    $rEf = 0.0; $rTr = 0.0
    if ($met -eq "Transferencia") { $rTr = $baseDev * 2 }
    elseif ($met -eq "Combinado") {
        $pe = 0.0; try { $pe = [double]$pedido.pagoEfectivo } catch {}
        $pt = 0.0; try { $pt = [double]$pedido.pagoTransferencia } catch {}
        if (($pe + $pt) -gt 0) { $rEf = $baseDev * ($pe / ($pe + $pt)); $rTr = $baseDev * ($pt / ($pe + $pt)) * 2 } else { $rEf = $baseDev }
    }
    else { $rEf = $baseDev }
    $rEf = [math]::Round($rEf, 2); $rTr = [math]::Round($rTr, 2)
    $partesMet = @()
    if ($rEf -gt 0) { $partesMet += ("Efectivo $" + (Formato-Monto $rEf)) }
    if ($rTr -gt 0) { $partesMet += ("Transferencia $" + (Formato-Monto $rTr)) }
    $txtMet = ($partesMet -join " + ")

    $reintegrar = $true
    try { if ($d.PSObject.Properties.Name -contains 'reintegrarStock') { $reintegrar = [bool]$d.reintegrarStock } } catch { $reintegrar = $true }
    if ($reintegrar) {
        foreach ($it in $itemsDev) {
            if ([string]::IsNullOrWhiteSpace($it.sku)) { continue }
            $prodD = $global:catalogo | Where-Object { $_.sku -eq $it.sku } | Select-Object -First 1
            if ($prodD -and $prodD.stockBase -ne $null) {
                $prodD.vendido = [double]$prodD.vendido - [double]$it.cantidad
                if ($prodD.vendido -lt 0) { $prodD.vendido = 0 }
                $prodD.stock = [double]$prodD.stockBase - [double]$prodD.vendido
                if ($prodD.stock -lt 0) { $prodD.stock = 0 }
            }
        }
    }

    $registro = [pscustomobject]@{
        id                     = $global:nextIdDevolucion
        tipo                   = "devolucion"
        hora                   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        pedidoId               = [string]$idPed
        pedidoHora             = [string]$pedido.hora
        cliente                = ""
        atendio                = $vendedor
        motivo                 = ([string]$d.motivo).Trim()
        items                  = @($itemsDev)
        total                  = $totalDev
        metodoReembolso        = $txtMet
        entregado              = ""
        diferencia             = $null
        reintegrarStock        = $reintegrar
        reembolsoEfectivo      = $rEf
        reembolsoTransferencia = $rTr
        origen                 = "movil"
    }
    $global:nextIdDevolucion++
    [void]$global:devoluciones.Add($registro)
    Guardar-Devoluciones

    $impreso = $false
    $errImp = ""
    try { $impreso = [bool](Imprimir-Texto (Generar-TextoDevolucion $registro)) } catch { $errImp = "$($_.Exception.Message)" }
    return @{ codigo = 200; cuerpo = @{ ok = $true; id = $registro.id; impreso = $impreso; errorImpresion = $errImp; reembolsoEfectivo = $rEf; reembolsoTransferencia = $rTr } }
}

# ------------------------------------------------------------------
# Mensajes cortos movil <-> PC (solo entre la caja y los vendedores, nunca clientes)
# ------------------------------------------------------------------
# - Movil -> PC: una nota opcional que viaja con el pedido (campo "nota" del pedido) o un
#   mensaje suelto sin productos. Los sueltos quedan en mensajes_pc.json hasta que el Panel
#   de la PC los muestra (si estaba cerrado, los ve al abrirlo).
# - PC -> movil: un mensaje suelto llega como aviso del movil (campana); si va con un pedido
#   armado para el vendedor, la nota aparece dentro de ese pedido.
$mensajesPcPath = Join-Path $scriptDir "mensajes_pc.json"
$global:mensajesPc = New-Object System.Collections.ArrayList
$global:nextIdMensaje = 1

function Cargar-MensajesPc {
    if (Test-Path $mensajesPcPath) {
        try {
            $raw = Get-Content $mensajesPcPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                foreach ($m in @($raw | ConvertFrom-Json)) {
                    if (-not $m) { continue }
                    [void]$global:mensajesPc.Add($m)
                    if ([int]$m.id -ge $global:nextIdMensaje) { $global:nextIdMensaje = [int]$m.id + 1 }
                }
            }
        } catch { Write-Host "Aviso: no se pudo leer mensajes_pc.json anterior." }
    }
}

function Guardar-MensajesPc {
    while ($global:mensajesPc.Count -gt 100) { $global:mensajesPc.RemoveAt(0) }
    Escribir-ArchivoConReintento -ruta $mensajesPcPath -contenido (ConvertTo-Json -InputObject @($global:mensajesPc) -Depth 4) | Out-Null
}

Cargar-MensajesPc

# Deja el texto en una sola linea y de hasta 200 caracteres.
function Limpiar-TextoMensaje($t) {
    $s = ([string]$t).Trim()
    $s = $s -replace '[\r\n\t]+', ' '
    if ($s.Length -gt 200) { $s = $s.Substring(0, 200) }
    return $s
}

# ------------------------------------------------------------------
# HTML: Panel Receptor (se abre en la PC)
# ------------------------------------------------------------------
$htmlPC = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Panel de Pedidos - Toto Tools</title>
<link rel="manifest" href="/manifest-pc.json">
<link rel="icon" href="/icon-pc-192.png">
<link rel="apple-touch-icon" href="/icon-pc-192.png">
<meta name="theme-color" content="#0f172a">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Panel Pedidos">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<style>
  * { box-sizing: border-box; margin:0; padding:0; font-family: 'Segoe UI', system-ui, Tahoma, Arial, sans-serif; }
  body { background:#0f172a; color:#e2e8f0; padding:18px; }
  header { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:14px; padding-bottom:14px; border-bottom:1px solid #1e293b; }
  h1 { font-size:21px; color:#fff; font-weight:600; }
  .share { background:#1e293b; padding:10px 14px; border-radius:8px; font-size:13px; color:#93c5fd; }
  .share b { color:#fff; }
  .catalogoInfo { font-size:12px; color:#94a3b8; margin-bottom:16px; display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
  .catalogoInfo button { background:#334155; color:#cbd5e1; border:none; padding:6px 10px; border-radius:6px; font-size:12px; cursor:pointer; }
  .catalogoInfo .error { color:#fca5a5; }
  .toolbar { display:flex; gap:10px; align-items:center; margin-bottom:18px; flex-wrap:wrap; }
  .toolbar label { font-size:14px; display:flex; align-items:center; gap:6px; }
  #banner { display:none; background:#16a34a; color:#fff; padding:12px 16px; border-radius:8px; margin-bottom:16px; font-weight:bold; text-align:center; box-shadow:0 4px 12px rgba(22,163,74,0.35); }
  #grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(270px,1fr)); gap:16px; }
  .card { border-radius:12px; padding:16px; color:#0f172a; box-shadow:0 3px 10px rgba(0,0,0,0.25); }
  .pendiente { background:#fde047; }
  .cobrado { background:#86efac; }
  .cancelado { background:#cbd5e1; opacity:0.75; }
  .porrevisar { background:#93c5fd; outline:3px solid #2563eb; }
  .revision { background:rgba(255,255,255,0.65); border-radius:8px; padding:8px 10px; margin:6px 0; font-size:13px; }
  .revision b { display:block; font-size:14px; margin-bottom:3px; }
  .revision-cuadre { display:flex; flex-wrap:wrap; gap:6px; margin:6px 0; }
  .revision-cuadre input { flex:1 1 100%; min-width:0; padding:8px; border-radius:7px; border:1px solid #cbd5e1; font-size:14px; }
  .revision-cuadre button { flex:1 1 auto; width:auto; white-space:nowrap; }
  .discrepancia-aviso { color:#991b1b; background:#fee2e2; border:1px solid #fca5a5; border-radius:8px; padding:8px; font-size:12px; margin:6px 0; }
  .btn-ok { background:#15803d; }
  .btn-ok:hover { background:#166534; }
  #contadorRevisar { display:none; background:#2563eb; color:#fff; padding:5px 12px; border-radius:12px; font-size:12px; font-weight:bold; }
  .card h3 { font-size:16px; margin-bottom:4px; }
  .card .hora { font-size:11px; opacity:0.7; margin-bottom:10px; }
  .card ul { list-style:none; font-size:13px; margin-bottom:8px; }
  .card ul li { padding:3px 0; border-bottom:1px dashed rgba(0,0,0,0.15); display:flex; align-items:center; gap:8px; }
  .card ul li .ped-miniatura { width:36px; height:36px; border-radius:6px; object-fit:cover; flex-shrink:0; cursor:zoom-in; background:rgba(0,0,0,0.08); }
  .card .total { font-weight:bold; font-size:16px; margin:4px 0; }
  .card .estado { display:inline-block; font-size:11px; font-weight:bold; padding:3px 9px; border-radius:12px; background:rgba(0,0,0,0.15); margin-bottom:8px; letter-spacing:0.5px; }
  .btn { background:#1d4ed8; color:#fff; border:none; padding:9px 12px; border-radius:8px; font-weight:bold; cursor:pointer; width:100%; font-size:14px; margin-top:4px; }
  .btn:hover { background:#1e40af; }
  .btn-imprimir { background:#475569; }
  .btn-imprimir:hover { background:#334155; }
  .btn-cancelar { background:#b91c1c; }
  .btn-cancelar:hover { background:#991b1b; }
  .btn-cerrar-dia { background:#7c2d12; color:#fff; border:none; padding:8px 14px; border-radius:8px; font-size:13px; font-weight:bold; cursor:pointer; }
  .btn-cerrar-dia:hover { background:#5c2109; }
  .btn-toggle { background:#334155; color:#cbd5e1; border:none; padding:8px 14px; border-radius:8px; font-size:13px; font-weight:bold; cursor:pointer; }
  .btn-toggle:hover { background:#475569; }
  .btn-toggle.activo { background:#16a34a; color:#fff; }
  .btn-toggle.activo:hover { background:#15803d; }
  #btnMenuPC { position:relative; }
  #btnMenuPC.alerta::after { content:''; position:absolute; top:2px; right:2px; width:10px; height:10px; border-radius:50%; background:#ef4444; }
  .vacio { text-align:center; color:#64748b; padding:40px; }
  .btn-nuevo-pedido { background:#7c3aed; color:#fff; border:none; padding:8px 14px; border-radius:8px; font-size:13px; font-weight:bold; cursor:pointer; }
  .btn-nuevo-pedido:hover { background:#6d28d9; }
  .np-select { width:100%; padding:9px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:14px; margin-bottom:10px; }
  .np-input { width:100%; padding:9px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:14px; margin-bottom:8px; }
  .np-resultado { display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid #334155; font-size:13px; }
  .np-resultado .np-miniatura { width:38px; height:38px; border-radius:6px; object-fit:cover; margin-right:8px; flex-shrink:0; cursor:zoom-in; background:#0f172a; }
  .np-resultado .np-info { flex:1; min-width:0; }
  .np-resultado .np-info b { display:block; }
  .np-resultado .np-info span { color:#94a3b8; font-size:12px; }
  .np-resultado button { background:#334155; color:#fff; border:none; width:32px; height:32px; border-radius:7px; font-size:17px; font-weight:bold; cursor:pointer; flex-shrink:0; }
  .np-carrito-item { display:flex; justify-content:space-between; align-items:center; padding:7px 0; border-bottom:1px dashed #334155; font-size:13px; }
  .np-carrito-item button { background:#334155; color:#fff; border:none; width:26px; height:26px; border-radius:6px; font-weight:bold; cursor:pointer; }
  .np-total { text-align:right; font-weight:bold; font-size:15px; margin:8px 0; }
  .np-vacio { color:#64748b; font-size:13px; }

  #tarjetaImprimir { display:none; }
  @media print {
    body * { visibility: hidden; }
    #tarjetaImprimir, #tarjetaImprimir * { visibility: visible; }
    #tarjetaImprimir {
      display: block; position: fixed; top:0; left:0; width:100%;
      background:#fff; color:#000; padding:24px;
    }
  }
</style>
</head>
<body>
  <header>
    <h1>Panel de Pedidos - Toto Tools</h1>
    <div class="share">Vendedores conectan en: <b id="shareUrl">...</b></div>
    <button id="btnStockBajo" class="btn-toggle" style="font-size:15px; position:relative;" onclick="togglePanelStockBajo()" title="Productos con stock bajo">&#128276;<span id="badgeStockBajo" style="display:none; position:absolute; top:-6px; right:-6px; min-width:18px; height:18px; padding:0 4px; border-radius:9px; background:#ef4444; color:#fff; font-size:11px; font-weight:bold; align-items:center; justify-content:center;">0</span></button>
    <button id="btnMenuPC" class="btn-toggle" style="font-size:15px;" onclick="abrirMenuPC()">&#9881; Ajustes</button>
  </header>

  <div class="catalogoInfo" id="catalogoInfo">Cargando catalogo...</div>

  <div class="toolbar">
    <label><input type="checkbox" id="soloPendientes" checked> Mostrar solo pendientes y por revisar</label>
    <span id="contadorRevisar"></span>
    <button class="btn-nuevo-pedido" onclick="abrirNuevoPedido()">Nuevo pedido para vendedor</button>
  </div>

  <div id="panelStockBajo" style="display:none; position:fixed; top:64px; right:16px; z-index:57; width:340px; max-width:92vw; max-height:70vh; overflow:auto; background:#1e293b; border:1px solid #334155; border-radius:12px; padding:12px; box-shadow:0 8px 24px rgba(0,0,0,0.5);">
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
      <strong style="color:#fff; font-size:14px;">Stock bajo</strong>
      <span><button class="btn-toggle" style="font-size:12px; padding:4px 8px; background:#0369a1; color:#fff;" onclick="abrirCompras()">Lista de compras</button> <button class="btn-toggle" style="font-size:12px; padding:4px 8px;" onclick="descartarTodoStockBajo()">Borrar todas</button> <button class="btn-toggle" style="font-size:12px; padding:4px 8px;" onclick="togglePanelStockBajo()">Cerrar</button></span>
    </div>
    <div id="listaStockBajo" style="font-size:13px; color:#e2e8f0;"></div>
  </div>
  <div id="banner">Nuevo pedido recibido</div>
  <div id="grid"></div>

  <div id="menuPCOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:58;" onclick="if (event.target === this) cerrarMenuPC()">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:420px; width:92%; max-height:88vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:14px; color:#fff;">Ajustes</h2>

      <div id="resumenHoyMenu" style="font-size:13px; color:#cbd5e1; background:#0f172a; border-radius:8px; padding:10px 12px; margin-bottom:16px;">Cargando resumen de hoy...</div>

      <button class="btn-toggle" style="width:100%; margin-bottom:10px; text-align:left;" onclick="cerrarMenuPC(); abrirConfig();">Cargar / cambiar el Excel del catalogo</button>

      <label style="font-size:14px; display:flex; align-items:center; gap:8px; margin-bottom:14px;" title="Se imprime como TOTAL USD en los recibos. Deja vacio o 0 para no imprimirlo.">Tasa USD: <input type="number" id="tasaDolar" min="0" step="any" placeholder="0" style="width:90px; padding:6px 8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0;" onchange="guardarTasaDolar()"></label>

      <label style="font-size:14px; display:flex; align-items:center; gap:8px; margin-bottom:14px;" title="Cuando un producto queda con esta cantidad o menos, avisa en el Panel y se marca en el buscador de vendedores y clientes. Deja 0 para desactivar el aviso.">Alerta stock bajo (unid.): <input type="number" id="umbralStockBajo" min="0" step="1" placeholder="3" style="width:70px; padding:6px 8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0;" onchange="guardarUmbralStockBajo()"></label>

      <button class="btn-toggle" id="btnPermitirDescuentos" style="width:100%; margin-bottom:10px;" onclick="alternarPermitirDescuentos()" title="Cuando esta activado, el vendedor puede editar el precio de cada linea de su pedido (descuento)">Descuentos por producto: ...</button>

      <button class="btn-toggle" id="btnOcultarSinStock" style="width:100%; margin-bottom:10px;" onclick="alternarOcultarSinStock()" title="Cuando esta activado, a los vendedores no les aparecen en el buscador los productos sin existencias">Ocultar sin stock a vendedores: ...</button>

      <button class="btn-toggle" id="btnAutoservicioDestino" style="width:100%; margin-bottom:10px;" onclick="alternarAutoservicioDestino()" title="Adonde caen los pedidos que arman los clientes ellos mismos desde el enlace de autoservicio">Pedidos de clientes (autoservicio): ...</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="copiarEnlaceAutoservicio()" title="Enlace para que los clientes vean el catalogo completo y armen su propio pedido desde su telefono">Copiar enlace para clientes</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="mostrarQRCliente()" title="Muestra un codigo QR con el enlace del catalogo para que los clientes lo escaneen con la camara del telefono">Ver QR para clientes</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="mostrarQRVendedor()" title="Muestra un codigo QR con el enlace de la app de vendedor, para que se enlacen escaneandolo con la camara sin escribir la direccion IP">Ver QR para vendedores</button>

      <div style="margin-bottom:10px; padding:10px; background:#0f172a; border-radius:8px; border:1px solid #334155;">
        <label style="font-size:12px; color:#94a3b8; display:block; margin-bottom:4px;">Nombre de la red WiFi (SSID)</label>
        <input type="text" id="inputWifiSSID" placeholder="Ej: TotoTools" style="width:100%; padding:8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:13px; margin-bottom:8px;">
        <label style="font-size:12px; color:#94a3b8; display:block; margin-bottom:4px;">Contraseña del WiFi</label>
        <input type="text" id="inputWifiClave" placeholder="Dejar vacio si la red es abierta" style="width:100%; padding:8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:13px; margin-bottom:8px;">
        <button class="btn-toggle" style="width:100%;" onclick="guardarYMostrarQRWifi()" title="El cliente escanea este QR y su telefono se conecta solo al WiFi, sin ver ni escribir la contraseña">Ver QR para conectarse al WiFi</button>
        <button class="btn-toggle" style="width:100%; margin-top:8px;" onclick="generarTarjetaImprimir()" title="Genera una hoja con los dos QR (WiFi y catalogo) lista para imprimir o mandar por WhatsApp">Generar tarjeta para imprimir (WiFi + Catalogo)</button>
      </div>

      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="window.open('/etiquetas', '_blank')">Etiquetas y códigos de barra</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="abrirCompras()">Lista de compras (reabastecer)</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="abrirDevolucion()">Ticket de devolución / garantía</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="abrirPermisos()">Permisos de vendedores</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="abrirRed()">Red / IP fija de la PC</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:10px;" onclick="window.open('/metricas', '_blank')">Ver metricas</button>

      <button class="btn-nuevo-pedido" style="width:100%; background:#0369a1; margin-bottom:10px;" onclick="cerrarMenuPC(); abrirPines();">PIN de vendedores</button>

      <button class="btn-toggle" id="btnFotosCatalogo" style="width:100%; margin-bottom:10px;" onclick="importarFotosCatalogo()" title="Lee el ultimo backup-catalogo-*.zip de esta carpeta (Copia de seguridad de la app del catalogo) y actualiza las fotos por SKU">Fotos del catalogo: ...</button>
      <button class="btn-toggle" style="width:100%; margin-bottom:14px;" onclick="document.getElementById('zipFotosPC').click()" title="Elige el .zip de Copia de seguridad del catalogo desde cualquier carpeta de esta PC">Subir .zip de fotos...</button>
      <input type="file" id="zipFotosPC" accept=".zip" style="display:none">

      <button class="btn-cerrar-dia" style="width:100%; margin-bottom:10px;" onclick="cerrarMenuPC(); cerrarDia();">Cerrar el dia (archivar cobrados)</button>

      <button class="btn" style="background:#475569;" onclick="cerrarMenuPC()">Cerrar</button>
    </div>
  </div>

  <div id="pinesOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:55;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:420px; width:92%; max-height:88vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:6px; color:#fff;">PIN de vendedores</h2>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:12px;">Un vendedor solo se puede crear desde aqui: nombre que no este repetido y un PIN de 4 a 6 numeros (obligatorio). Sin estar creado, nadie puede entrar como ese vendedor desde el movil.</p>
      <label style="font-size:13px; display:block; margin-bottom:4px; color:#cbd5e1;">Nombre del vendedor:</label>
      <input type="text" id="pinVendedorNombre" class="np-input" placeholder="Ej: Jose">
      <label style="font-size:13px; display:block; margin-bottom:4px; color:#cbd5e1;">PIN (4 a 6 numeros, obligatorio):</label>
      <input type="text" id="pinVendedorValor" class="np-input" inputmode="numeric" placeholder="Ej: 1234">
      <button class="btn" onclick="guardarPin()">Guardar</button>
      <div id="pinError" style="color:#fca5a5; font-size:12px; margin-top:8px;"></div>
      <h3 style="font-size:13px; color:#94a3b8; margin:14px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Vendedores con PIN puesto</h3>
      <div id="listaConPin" style="font-size:13px; color:#e2e8f0;">Cargando...</div>
      <h3 style="font-size:13px; color:#94a3b8; margin:14px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Modo punto de venta en el m&oacute;vil</h3>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:8px;">Con esta clave de administrador, un vendedor puede activar en su m&oacute;vil (Ajustes) el modo punto de venta: historial de ventas, devoluciones, descuentos y reporte de efectivo y transferencia. Si la cambias o la quitas, todos los m&oacute;viles pierden el modo.</p>
      <input type="text" id="posClaveAdmin" class="np-input" placeholder="Clave de administrador (4 a 20 caracteres)" autocomplete="off">
      <button class="btn" onclick="guardarClavePOS()">Guardar clave</button>
      <div id="posClaveEstado" style="font-size:12px; color:#94a3b8; margin-top:6px;"></div>
      <button class="btn" style="background:#475569; margin-top:12px;" onclick="cerrarPines()">Cerrar</button>
    </div>
  </div>

  <div id="comprasOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:59;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:560px; width:94%; max-height:90vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:4px; color:#fff;">Lista de compras (reabastecer)</h2>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:10px;">Productos que llegaron a su mínimo de seguridad. Ajusta cuánto comprar y desmarca lo que no quieras pedir.</p>
      <div id="comprasLista" style="max-height:38vh; overflow:auto; border:1px solid #334155; border-radius:8px; padding:4px 8px;"></div>
      <div id="comprasResumen" style="font-size:12px; color:#94a3b8; margin:8px 0;"></div>
      <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:6px;">
        <button class="btn" style="flex:1; width:auto;" onclick="imprimirCompras()">Imprimir ticket</button>
        <button class="btn" style="flex:1; width:auto; background:#0369a1;" onclick="copiarCompras()">Copiar texto</button>
        <button class="btn" style="flex:1; width:auto; background:#475569;" onclick="hojaCompras()">Hoja / PDF</button>
      </div>
      <div id="comprasMsg" style="font-size:12px; color:#86efac; min-height:16px;"></div>
      <h3 style="font-size:13px; color:#94a3b8; margin:12px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Mínimo por producto</h3>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:6px;">Por defecto se usa el mínimo general de Ajustes. Aquí puedes fijar uno propio (0 = sin alerta para ese producto).</p>
      <input type="text" id="minBuscador" class="np-input" placeholder="Buscar producto (nombre o SKU)" oninput="buscarMinimo()">
      <div id="minResultados"></div>
      <button class="btn" style="background:#475569; margin-top:12px;" onclick="cerrarCompras()">Cerrar</button>
    </div>
  </div>

  <div id="devOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:59;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:560px; width:94%; max-height:92vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:10px; color:#fff;">Devolución / cambio / garantía</h2>
      <div style="display:flex; gap:8px;">
        <select id="devTipo" class="np-select" style="flex:1;" onchange="devTipoCambio()">
          <option value="devolucion">Devolución</option>
          <option value="cambio">Cambio de producto</option>
          <option value="garantia">Garantía</option>
        </select>
        <input type="number" id="devPedido" class="np-input" style="flex:1;" placeholder="Pedido # (opcional)">
        <button class="btn-toggle" style="height:38px;" onclick="devCargarPedido()">Cargar</button>
      </div>
      <input type="text" id="devCliente" class="np-input" placeholder="Cliente (opcional)">
      <input type="text" id="devBuscar" class="np-input" placeholder="Agregar producto recibido (nombre o SKU)" oninput="devBuscarProd()">
      <div id="devResultados"></div>
      <h3 style="font-size:13px; color:#94a3b8; margin:10px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Productos recibidos</h3>
      <div id="devLista"><div class="np-vacio">Sin productos aún.</div></div>
      <div class="np-total" style="margin:6px 0;">Total: $<span id="devTotal">0.00</span></div>
      <textarea id="devMotivo" class="np-input" rows="2" placeholder="Motivo (ej: no enciende, defecto de fábrica)"></textarea>
      <div style="display:flex; gap:8px;">
        <select id="devMetodo" class="np-select" style="flex:1;">
          <option value="Efectivo">Reembolso en efectivo</option>
          <option value="Transferencia">Reembolso por transferencia</option>
          <option value="Cambio de producto">Cambio de producto</option>
          <option value="Nota de credito">Nota de crédito</option>
          <option value="">Sin reembolso</option>
        </select>
        <input type="number" id="devDif" class="np-input" style="flex:1;" placeholder="Diferencia ± (opcional)">
      </div>
      <input type="text" id="devEntregado" class="np-input" placeholder="Entregado en cambio (texto, opcional)">
      <label style="font-size:13px; display:flex; align-items:center; gap:8px; margin-bottom:10px; color:#cbd5e1;"><input type="checkbox" id="devStock" checked> Devolver estas unidades al stock</label>
      <button class="btn" onclick="devGuardar()">Guardar e imprimir comprobante</button>
      <div id="devMsg" style="font-size:12px; margin-top:8px; min-height:16px;"></div>
      <h3 style="font-size:13px; color:#94a3b8; margin:12px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Últimos comprobantes</h3>
      <div id="devHistorial" style="font-size:13px;"></div>
      <button class="btn" style="background:#475569; margin-top:12px;" onclick="cerrarDevolucion()">Cerrar</button>
    </div>
  </div>

  <div id="permOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:59;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:480px; width:94%; max-height:92vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:4px; color:#fff;">Permisos de vendedores</h2>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:10px;">Decide qué puede hacer y a qué partes puede entrar cada vendedor. Los cambios se guardan al momento y el móvil los aplica en unos segundos (se le avisa).</p>
      <select id="permVendedor" class="np-select" onchange="permSeleccionar(this.value)"></select>
      <div style="display:flex; gap:6px; flex-wrap:wrap; margin-bottom:10px;">
        <button class="btn-toggle" style="font-size:12px; padding:6px 10px;" onclick="permPreset('completo')">Completo</button>
        <button class="btn-toggle" style="font-size:12px; padding:6px 10px;" onclick="permPreset('vender')">Solo vender</button>
        <button class="btn-toggle" style="font-size:12px; padding:6px 10px;" onclick="permPreset('consulta')">Solo consulta</button>
      </div>
      <div id="permLista"></div>
      <div id="permMsg" style="font-size:12px; color:#86efac; min-height:16px; margin-top:6px;"></div>
      <button class="btn" style="background:#475569; margin-top:10px;" onclick="cerrarPermisos()">Cerrar</button>
    </div>
  </div>

  <div id="redOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:59;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:460px; width:94%; max-height:92vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:4px; color:#fff;">Red / IP fija de la PC</h2>
      <p style="font-size:12px; color:#94a3b8; margin-bottom:10px;">Cada punto de acceso da otro rango de IP. Aquí eliges el último número (ej. 50) y la PC toma siempre ese número en la red donde esté conectada.</p>
      <div id="redInfo" style="font-size:13px; background:#0f172a; border-radius:8px; padding:10px 12px; margin-bottom:10px; color:#cbd5e1;">Cargando...</div>
      <label style="font-size:13px; color:#cbd5e1; display:block; margin-bottom:4px;">Último número de la IP (2 a 254):</label>
      <input type="number" id="redOcteto" class="np-input" min="2" max="254" placeholder="Ej: 50">
      <label style="font-size:13px; display:flex; align-items:center; gap:8px; margin-bottom:10px; color:#cbd5e1;"><input type="checkbox" id="redAuto"> Aplicarlo solo al abrir el servidor (en la red donde esté)</label>
      <button class="btn" onclick="fijarIpRed()">Fijar IP en esta red</button>
      <button class="btn" style="background:#475569; margin-top:6px;" onclick="ipAutomaticaRed()">Volver a IP automática (DHCP)</button>
      <div id="redMsg" style="font-size:12px; margin-top:8px; min-height:16px; line-height:1.4;"></div>
      <button class="btn" style="background:#334155; margin-top:10px;" onclick="pcById('redOverlay').style.display='none'">Cerrar</button>
    </div>
  </div>

  <div id="nuevoPedidoOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:55;">
    <div style="background:#1e293b; padding:22px; border-radius:12px; max-width:460px; width:92%; max-height:88vh; overflow:auto;">
      <h2 id="npTitulo" style="font-size:17px; margin-bottom:12px; color:#fff;">Nuevo pedido para vendedor</h2>
      <div id="npDestinoWrap">
        <label style="font-size:13px; display:block; margin-bottom:4px; color:#cbd5e1;">Enviar a:</label>
        <select id="selVendedorDestino" class="np-select"></select>
        <div id="npSinVendedores" style="display:none; color:#fca5a5; font-size:12px; margin:-4px 0 8px;">Nadie est&aacute; en l&iacute;nea ahora mismo: el pedido queda guardado y les aparece cuando abran la app, siempre que sea dentro del plazo de abajo.</div>
      </div>
      <div id="npVigenciaWrap">
        <label style="font-size:13px; display:block; margin-bottom:4px; color:#cbd5e1;">El pedido espera a los vendedores durante:</label>
        <select id="selVigenciaNP" class="np-select"><option value="30">30 minutos</option><option value="60" selected>1 hora</option><option value="120">2 horas</option><option value="480">8 horas</option></select>
      </div>
      <label style="font-size:13px; display:block; margin-bottom:4px; color:#cbd5e1;">Buscar producto:</label>
      <input type="text" id="npBuscador" class="np-input" placeholder="Nombre o SKU" oninput="buscarNP()">
      <div id="npResultados"></div>
      <h3 style="font-size:13px; color:#94a3b8; margin:12px 0 6px; text-transform:uppercase; letter-spacing:0.5px;">Pedido a armar</h3>
      <div id="npCarrito"><div class="np-vacio">Sin productos aun.</div></div>
      <div class="np-total">Total: $<span id="npTotal">0.00</span></div>
      <div id="npNotaWrap"><input type="text" id="npNota" class="np-input" maxlength="200" placeholder="Mensaje para el vendedor (opcional; sin productos se env&iacute;a solo el mensaje)"></div>
      <button class="btn" id="npBtnEnviar" onclick="enviarPedidoAVendedor()" style="margin-top:6px;">Enviar al movil del vendedor</button>
      <button class="btn" style="background:#475569; margin-top:6px;" onclick="cerrarNuevoPedido()">Cancelar</button>
      <div id="npError" style="color:#fca5a5; font-size:12px; margin-top:8px;"></div>
    </div>
  </div>

  <div id="fotoOverlayNP" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.85); align-items:center; justify-content:center; z-index:60;" onclick="cerrarFotoProductoNP()">
    <div style="text-align:center;" onclick="event.stopPropagation()">
      <img id="fotoOverlayImgNP" src="" style="max-width:94vw; max-height:80vh; border-radius:10px; background:#fff;">
      <div id="fotoOverlayNombreNP" style="color:#fff; margin-top:10px; font-size:15px;"></div>
    </div>
  </div>

  <div id="configOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:50;">
    <div style="background:#1e293b; padding:24px; border-radius:12px; max-width:480px; width:92%; max-height:85vh; overflow:auto;">
      <h2 style="font-size:17px; margin-bottom:14px; color:#fff;">Configurar Excel del catalogo</h2>

      <div id="configPaso1">
        <label style="font-size:13px; display:block; margin-bottom:6px;">Ruta completa del archivo Excel en esta PC:</label>
        <input type="text" id="inputRuta" placeholder="C:\ToTo Tools\catalogo.xlsx" style="width:100%; padding:9px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:13px; margin-bottom:10px;">
        <button class="btn" onclick="guardarRutaYContinuar()">Continuar</button>
        <div id="configPaso1Error" style="color:#fca5a5; font-size:12px; margin-top:8px;"></div>
      </div>

      <div id="configPaso2" style="display:none;">
        <p style="font-size:12px; color:#94a3b8; margin-bottom:10px;">Elige que columna del Excel corresponde a cada dato (SKU y Cantidad/Stock son opcionales):</p>
        <div id="mapeoCampos"></div>
        <button class="btn" onclick="guardarMapeoYCerrar()" style="margin-top:10px;">Guardar</button>
        <button class="btn" style="background:#475569; margin-top:6px;" onclick="cerrarConfig()">Cancelar</button>
        <div id="configPaso2Error" style="color:#fca5a5; font-size:12px; margin-top:8px;"></div>
      </div>
    </div>
  </div>

  <div id="qrOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:50;" onclick="cerrarQRCliente()">
    <div style="background:#1e293b; padding:24px; border-radius:12px; max-width:340px; width:88%; text-align:center;" onclick="event.stopPropagation()">
      <h2 id="qrTituloModal" style="font-size:16px; margin-bottom:14px; color:#fff;">QR para clientes</h2>
      <div id="qrCodeBox" style="background:#fff; padding:12px; border-radius:8px; display:inline-block; min-width:200px; min-height:200px;"></div>
      <div id="qrEnlaceTexto" style="font-size:11px; color:#94a3b8; margin-top:12px; word-break:break-all;"></div>
      <button id="qrBotonCopiar" class="btn" style="margin-top:14px; width:100%;" onclick="copiarEnlaceQRActual()">Copiar enlace</button>
      <button class="btn" style="background:#475569; margin-top:8px; width:100%;" onclick="cerrarQRCliente()">Cerrar</button>
    </div>
  </div>

  <div id="tarjetaImprimir">
    <h2 style="text-align:center; margin-bottom:20px;">Toto Tools</h2>
    <div style="display:flex; justify-content:space-around; gap:16px; flex-wrap:wrap;">
      <div style="text-align:center; max-width:45%;">
        <p style="font-weight:bold; margin-bottom:8px;">1. Conectate al WiFi</p>
        <div id="tarjetaQRWifi" style="display:inline-block;"></div>
        <p id="tarjetaWifiNombre" style="font-size:12px; margin-top:6px;"></p>
      </div>
      <div style="text-align:center; max-width:45%;">
        <p style="font-weight:bold; margin-bottom:8px;">2. Mira el catalogo</p>
        <div id="tarjetaQRCatalogo" style="display:inline-block;"></div>
      </div>
    </div>
  </div>

  <script>
    // Libreria QR (qrcode-generator de kazuhikoarase, MIT, sin dependencias
    // externas -- va incrustada aqui mismo, el navegador no descarga nada).
    var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())};return f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,(function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1}))},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(n*r,o*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default;var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();
  </script>
  <script src="/xlsx.js"></script>
  <script>
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {});
    }
    let ultimoMaxId = 0;
    let idsRevisarAnt = new Set();
    let primerCarga = true;
    let configMapeoEncabezados = [];
    let procesandoExcel = false;

    document.getElementById('shareUrl').textContent = location.origin + '/vendedor';

    function beep() {
      try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();
        const tocar = (inicio) => {
          const o = ctx.createOscillator();
          const g = ctx.createGain();
          o.connect(g); g.connect(ctx.destination);
          o.type = 'square';
          o.frequency.setValueAtTime(880, ctx.currentTime + inicio);
          g.gain.setValueAtTime(0.0001, ctx.currentTime + inicio);
          g.gain.exponentialRampToValueAtTime(0.9, ctx.currentTime + inicio + 0.02);
          o.start(ctx.currentTime + inicio);
          o.frequency.setValueAtTime(1200, ctx.currentTime + inicio + 0.15);
          g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + inicio + 0.5);
          o.stop(ctx.currentTime + inicio + 0.5);
        };
        tocar(0);
        tocar(0.55);
      } catch(e) {}
    }

    function mostrarBanner(texto, ms) {
      const b = document.getElementById('banner');
      if (bannerOcultarTimeout) { clearTimeout(bannerOcultarTimeout); bannerOcultarTimeout = null; }
      b.textContent = texto;
      b.style.display = 'block';
      bannerOcultarTimeout = setTimeout(() => { b.style.display = 'none'; }, ms || 3500);
    }

    // Igual que mostrarBanner, pero con un boton "Deshacer" adentro (mismo
    // recuadro, sin agregar nada nuevo a la pantalla). Se usa justo despues
    // de armar un pedido para vendedor(es), y tambien para el aviso de un
    // "para todos" que nadie tomo: en ambos casos "Deshacer" retira el
    // pedido, siempre que todavia nadie lo haya tomado.
    let bannerOcultarTimeout = null;
    function mostrarBannerDeshacer(texto, idAsignado) {
      const b = document.getElementById('banner');
      if (bannerOcultarTimeout) { clearTimeout(bannerOcultarTimeout); bannerOcultarTimeout = null; }
      b.innerHTML = '';
      b.appendChild(document.createTextNode(texto + '  '));
      const btn = document.createElement('button');
      btn.textContent = 'Deshacer';
      btn.style.cssText = 'background:#fff; color:#0369a1; border:none; border-radius:6px; padding:4px 10px; font-weight:700; cursor:pointer; margin-left:6px;';
      btn.onclick = () => deshacerAsignado(idAsignado);
      b.appendChild(btn);
      b.style.display = 'block';
      bannerOcultarTimeout = setTimeout(() => { b.style.display = 'none'; b.innerHTML = ''; }, 8000);
    }

    async function deshacerAsignado(id) {
      document.getElementById('banner').style.display = 'none';
      try {
        const res = await fetch('/api/pedidos/asignados/' + id + '/retirar', { method: 'POST' });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          mostrarBanner((data && data.error) || 'No se pudo deshacer (puede que ya lo hayan tomado).');
          return;
        }
        mostrarBanner('Pedido retirado.');
      } catch (e) { mostrarBanner('No se pudo deshacer (revisa la conexion).'); }
    }

    async function cobrar(id, metodoEsperado, totalCobrado) {
      // La caja YA NO escribe cuanto dio el cliente: eso solo lo pone el
      // vendedor cuando el mismo cobra un pedido pendiente desde su movil
      // (para que el cambio en efectivo salga bien en su recibo). Aqui la
      // caja solo confirma que cobro el pedido.
      const totalTxt = (typeof totalCobrado === 'number' && !isNaN(totalCobrado)) ? totalCobrado.toFixed(2) : '';
     try {
        // La PC (caja) no elige el metodo de pago: cobra con el que el
        // vendedor dejo en el pedido (Efectivo o Transferencia x2). Manda el
        // que tiene en pantalla para que el servidor avise si el vendedor
        // lo cambio justo antes.
        const res = await fetch('/api/pedidos/' + id + '/cobrar', { method:'POST', body: JSON.stringify({ metodoEsperado }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) alert((data && data.error) || 'No se pudo cobrar el pedido.');
        cargarPedidos();
      } catch(e) { alert('No se pudo actualizar el pedido.'); }
    }

    async function imprimir(id) {
      try {
        const res = await fetch('/api/pedidos/' + id + '/imprimir', { method:'POST' });
        const data = await res.json();
        if (!data.ok) { alert('No se pudo imprimir: ' + (data.error || 'error desconocido')); }
      } catch(e) {
        alert('No se pudo imprimir. Revisa el nombre de la impresora en el script.');
      }
    }

    // ------------------------------------------------------------
    // Configuracion del Excel (ruta + mapeo de columnas), una vez
    // ------------------------------------------------------------
    function abrirMenuPC() {
      document.getElementById('menuPCOverlay').style.display = 'flex';
      cargarResumenHoyMenu();
    }
    function cerrarMenuPC() { document.getElementById('menuPCOverlay').style.display = 'none'; }

    async function cargarResumenHoyMenu() {
      const el = document.getElementById('resumenHoyMenu');
      el.textContent = 'Cargando resumen de hoy...';
      try {
        const res = await fetch('/api/metricas?dias=1');
        const m = await res.json();
        if (!m.pedidosCobrados) { el.textContent = 'Hoy todavia no hay pedidos cobrados.'; return; }
        const porMetodo = (m.porMetodo || []).map(x => x.metodo + ': $' + x.monto.toFixed(2)).join(' · ');
        el.innerHTML = '<b>Hoy:</b> ' + m.pedidosCobrados + ' pedido(s) cobrados · Total: $' + m.totalFacturado.toFixed(2) +
          (porMetodo ? ('<br>' + porMetodo) : '');
      } catch (e) {
        el.textContent = 'No se pudo cargar el resumen de hoy.';
      }
    }

    function abrirConfig() {
      document.getElementById('configOverlay').style.display = 'flex';
      document.getElementById('configPaso1').style.display = 'block';
      document.getElementById('configPaso2').style.display = 'none';
      document.getElementById('configPaso1Error').textContent = '';
      fetch('/api/catalogo/config').then(r => r.json()).then(cfg => {
        document.getElementById('inputRuta').value = (cfg && cfg.ruta) ? cfg.ruta : '';
      }).catch(() => {});
    }

    function cerrarConfig() {
      document.getElementById('configOverlay').style.display = 'none';
    }

    async function guardarRutaYContinuar() {
      const ruta = document.getElementById('inputRuta').value.trim();
      const errEl = document.getElementById('configPaso1Error');
      errEl.textContent = '';
      if (!ruta) { errEl.textContent = 'Escribe la ruta completa del archivo.'; return; }
      try {
        await fetch('/api/catalogo/config', { method: 'POST', body: JSON.stringify({ ruta }) });
        const res = await fetch('/api/catalogo/archivo');
        if (!res.ok) { errEl.textContent = 'No se encontro el archivo en esa ruta. Revisala e intenta de nuevo.'; return; }
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });
        if (!filas.length) { errEl.textContent = 'El archivo esta vacio.'; return; }
        filaEncabezado = filaEncabezadoPara(filas, null);
        configMapeoEncabezados = nombresColumnas(filas[filaEncabezado]);
        configMapeoFilasDatos = filas.slice(filaEncabezado + 1, filaEncabezado + 61);
        try { filaBaseExcel = XLSX.utils.decode_range(hoja['!ref']).s.r; } catch (e) { filaBaseExcel = 0; }
        mostrarPaso2();
      } catch (e) {
        errEl.textContent = 'No se pudo leer el archivo: ' + e.message;
      }
    }

    // ---- Fila de encabezados y columnas del Excel ----
    // Muchos Excel traen filas de titulo arriba (nombre del negocio, fecha,
    // logo...) y los encabezados reales unas filas mas abajo. En vez de
    // asumir que estan en la fila 1, se busca la fila que mas parece un
    // encabezado. Las columnas sin encabezado se llaman "Columna N" y la
    // ventana de columnas muestra un ejemplo de cada una para elegir facil.
    let filaEncabezado = 0;
    let filaBaseExcel = 0;   // fila real de Excel donde empieza el rango usado (algunas hojas empiezan en B2, etc.)
    let configMapeoFilasDatos = [];

    function textoCelda(v) { return String(v == null ? '' : v); }

    function escaparHtml(t) {
      return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // Nombres de columna de una fila de encabezados: sin encabezado ->
    // "Columna N"; si dos se llaman igual, la segunda lleva " (2)".
    function nombresColumnas(fila) {
      const usados = new Map();
      return (fila || []).map((c, i) => {
        const n = textoCelda(c).trim() === '' ? ('Columna ' + (i + 1)) : textoCelda(c);
        const veces = (usados.get(n) || 0) + 1;
        usados.set(n, veces);
        return veces > 1 ? (n + ' (' + veces + ')') : n;
      });
    }

    function detectarFilaEncabezado(filas) {
      const clave = /sku|codigo|código|code|nombre|name|producto|descrip|articulo|artículo|precio|price|stock|cantidad|quantity|existencia|disponib|inventario|unidad/;
      let mejor = 0, mejorPuntos = -1;
      const limite = Math.min(filas.length, 40);
      for (let r = 0; r < limite; r++) {
        let textos = 0, claves = 0;
        for (const c of (filas[r] || [])) {
          const t = textoCelda(c).trim();
          if (!t) continue;
          if (isNaN(Number(t.replace(',', '.')))) textos++;
          if (clave.test(t.toLowerCase())) claves++;
        }
        const puntos = claves * 10 + textos;
        if (puntos > mejorPuntos) { mejorPuntos = puntos; mejor = r; }
      }
      return mejor;
    }

    // Con un mapeo ya guardado: la fila que contiene las columnas guardadas
    // de Nombre y Precio. Si no aparece (otro archivo), se detecta sola.
    function filaEncabezadoPara(filas, mapeo) {
      if (mapeo && mapeo.nombre && mapeo.precio) {
        const limite = Math.min(filas.length, 60);
        for (let r = 0; r < limite; r++) {
          const celdas = nombresColumnas(filas[r]);
          if (celdas.includes(mapeo.nombre) && celdas.includes(mapeo.precio)) return r;
        }
      }
      return detectarFilaEncabezado(filas);
    }

    // Primer valor no vacio de la columna (ejemplo para la lista)
    function ejemploColumna(i) {
      for (const f of configMapeoFilasDatos) {
        const v = textoCelda(f[i]).trim();
        if (v !== '') return v.length > 26 ? (v.slice(0, 26) + '…') : v;
      }
      return '';
    }

    function estadisticasColumna(i) {
      const vistos = new Set();
      let total = 0, textos = 0, enteros = 0;
      for (const f of configMapeoFilasDatos) {
        const t = textoCelda(f[i]).trim();
        if (t === '') continue;
        total++;
        vistos.add(t);
        if (isNaN(Number(t.replace(',', '.')))) textos++;
        else if (/^\d+$/.test(t)) enteros++;
      }
      return { total, textos, enteros, unicos: vistos.size };
    }

    // Cuando el encabezado no dice nada (columna sin titulo), se adivina por
    // lo que contiene: 'nombre' = texto casi siempre distinto; 'sku' = codigos
    // enteros casi todos distintos.
    function adivinarPorContenido(tipo, yaUsadas) {
      let mejor = -1, mejorPuntos = 0;
      for (let i = 0; i < configMapeoEncabezados.length; i++) {
        if (yaUsadas.includes(i)) continue;
        const e = estadisticasColumna(i);
        if (e.total < 3) continue;
        let puntos = 0;
        if (tipo === 'nombre') puntos = (e.textos / e.total > 0.8) ? e.unicos : 0;
        else puntos = (e.enteros / e.total > 0.9 && e.unicos / e.total > 0.9) ? e.unicos : 0;
        if (puntos > mejorPuntos) { mejorPuntos = puntos; mejor = i; }
      }
      return mejor;
    }

    // Prueba varias expresiones en orden de prioridad (la primera que
    // encuentre alguna columna gana), saltando las columnas "excluir".
    function adivinarColumna(regexes, excluir) {
      const lista = Array.isArray(regexes) ? regexes : [regexes];
      for (const rx of lista) {
        for (let i = 0; i < configMapeoEncabezados.length; i++) {
          const h = configMapeoEncabezados[i].toLowerCase().trim();
          if (!h || /^columna \d+$/.test(h)) continue;
          if (excluir && excluir.test(h)) continue;
          if (rx.test(h)) return i;
        }
      }
      return -1;
    }

    function preseleccionarColumnas() {
      const p = {
        precio: adivinarColumna(
          [/^\s*(precio\s*(de\s*)?venta|precio|pvp|retail\s*price|price)\s*$/, /precio\s*(de\s*)?venta|retail|pvp/, /precio|price/],
          /group|grupo|purchase|compra|wholesale|mayor|cost|\b[2-9]\b/),
        stock: adivinarColumna(
          [/^\s*(stock|cantidad|existencias?|disponible|final|available\s*quantity)\s*$/, /stock|cantidad|existencia|disponib|quantity|inventario/],
          /m[ií]n|nominal|group|grupo/),
        nombre: adivinarColumna(
          [/^\s*(nombre|name|producto|descripci[oó]n|art[ií]culo)\s*$/, /nombre|descrip|art[ií]culo|producto|\bname\b/],
          /c[oó]digo|code|sku|\bcod\b|\bref|group|grupo|catalog|\b[2-9]\b/),
        sku: adivinarColumna(
          [/^\s*(sku|c[oó]digo|code|cod|ref(erencia)?)\s*$/, /sku|c[oó]digo|\bcode\b|\bcod\b|\bref/],
          /barcode|bar\s*code|barras|catalog|group|grupo/)
      };
      const usadas = () => [p.precio, p.stock, p.nombre, p.sku].filter(i => i >= 0);
      if (p.nombre < 0) p.nombre = adivinarPorContenido('nombre', usadas());
      if (p.sku < 0) p.sku = adivinarPorContenido('sku', usadas());
      return p;
    }

    function mostrarPaso2() {
      document.getElementById('configPaso1').style.display = 'none';
      document.getElementById('configPaso2').style.display = 'block';

      const pre = preseleccionarColumnas();
      const campos = [
        { id: 'sku', etiqueta: 'SKU / Codigo', preseleccion: pre.sku, opcional: true },
        { id: 'nombre', etiqueta: 'Nombre del producto', preseleccion: pre.nombre, opcional: false },
        { id: 'precio', etiqueta: 'Precio', preseleccion: pre.precio, opcional: false },
        { id: 'stock', etiqueta: 'Cantidad / Stock', preseleccion: pre.stock, opcional: true }
      ];

      const cont = document.getElementById('mapeoCampos');
      cont.innerHTML = '<div style="font-size:12px; color:#94a3b8; margin-bottom:10px;">Encabezados detectados en la fila ' + (filaEncabezado + filaBaseExcel + 1) + ' del Excel.</div>' + campos.map(c => {
        const opciones = ['<option value="">' + (c.opcional ? '— Ninguna —' : '— elegir —') + '</option>']
          .concat(configMapeoEncabezados.map((h, i) => {
            const ej = ejemploColumna(i);
            if (!ej && /^Columna \d+$/.test(h)) return '';   // columna vacia y sin titulo
            return '<option value="' + i + '"' + (i === c.preseleccion ? ' selected' : '') + '>' + escaparHtml(h) + (ej ? '  —  ej: ' + escaparHtml(ej) : '') + '</option>';
          }))
          .join('');
        return '<div style="margin-bottom:10px;">' +
          '<label style="font-size:12px; display:block; margin-bottom:4px;">' + c.etiqueta + '</label>' +
          '<select id="mapeo-' + c.id + '" style="width:100%; padding:8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:13px;">' + opciones + '</select>' +
        '</div>';
      }).join('');
    }

    async function guardarMapeoYCerrar() {
      const errEl = document.getElementById('configPaso2Error');
      errEl.textContent = '';
      const val = id => {
        const sel = document.getElementById('mapeo-' + id);
        const v = sel.value;
        return v === '' ? null : configMapeoEncabezados[parseInt(v, 10)];
      };
      const mapeo = { sku: val('sku'), nombre: val('nombre'), precio: val('precio'), stock: val('stock') };
      if (!mapeo.nombre || !mapeo.precio) {
        errEl.textContent = 'Elige al menos la columna de Nombre y la de Precio.';
        return;
      }
      await fetch('/api/catalogo/config', { method: 'POST', body: JSON.stringify({ mapeo }) });
      cerrarConfig();
      await procesarExcel();
      cargarEstadoCatalogo();
    }

    // ------------------------------------------------------------
    // Lectura del Excel con SheetJS + envio del catalogo al servidor
    // ------------------------------------------------------------
    async function procesarExcel() {
      if (procesandoExcel) return;
      procesandoExcel = true;
      try {
        const resCfg = await fetch('/api/catalogo/config');
        const cfg = await resCfg.json();
        if (!cfg || !cfg.mapeo || !cfg.mapeo.nombre) { return; }

        const res = await fetch('/api/catalogo/archivo');
        if (!res.ok) throw new Error('no se pudo leer el archivo');
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });
        if (filas.length < 2) throw new Error('el archivo no tiene datos');

        const filaEnc = filaEncabezadoPara(filas, cfg.mapeo);
        const encabezados = nombresColumnas(filas[filaEnc]);
        const idx = nombreCol => nombreCol ? encabezados.indexOf(nombreCol) : -1;
        const iSku = idx(cfg.mapeo.sku);
        const iNombre = idx(cfg.mapeo.nombre);
        const iPrecio = idx(cfg.mapeo.precio);
        const iStock = idx(cfg.mapeo.stock);

        const limpiarNumero = v => {
          const n = parseFloat(String(v == null ? '' : v).replace(',', '.').replace(/[^0-9.\-]/g, ''));
          return isNaN(n) ? 0 : n;
        };

        const productos = [];
        for (let r = filaEnc + 1; r < filas.length; r++) {
          const fila = filas[r];
          const nombre = String(iNombre >= 0 ? (fila[iNombre] ?? '') : '').trim();
          if (!nombre) continue;
          const sku = iSku >= 0 ? String(fila[iSku] ?? '').trim() : '';
          const precio = iPrecio >= 0 ? limpiarNumero(fila[iPrecio]) : 0;
          let stock = null;
          if (iStock >= 0) {
            const v = fila[iStock];
            if (v !== '' && v !== null && v !== undefined) stock = limpiarNumero(v);
          }
          productos.push({ sku, nombre, precio: Math.round(precio * 100) / 100, stock });
        }

        await fetch('/api/catalogo/importar', { method: 'POST', body: JSON.stringify({ productos }) });
      } catch (e) {
        console.error('Error procesando Excel:', e);
      } finally {
        procesandoExcel = false;
      }
    }

    function renderCatalogoInfo(info) {
      const el = document.getElementById('catalogoInfo');
      if (!info.mapeoConfigurado) {
        el.innerHTML = '<span>Todavia no configuraste el Excel del catalogo.</span> <button onclick="abrirConfig()">Configurar Excel</button>';
        catalogoConError = true;
      } else if (info.error) {
        el.innerHTML = '<span class="error">Catalogo: ' + info.error + '</span> <button onclick="recargarCatalogo()">Reintentar</button> <button onclick="abrirConfig()">Reconfigurar</button>';
        catalogoConError = true;
      } else {
        el.innerHTML = 'Catalogo: ' + info.cantidad + ' productos (actualizado ' + (info.ultimaCarga || '-') + ') ' +
          '<button onclick="recargarCatalogo()">Recargar ahora</button> ' +
          '<button onclick="abrirConfig()">Reconfigurar columnas</button>';
        catalogoConError = false;
      }
      actualizarAlertaMenuPC();
    }

    async function recargarCatalogo() {
      try { await fetch('/api/catalogo/recargar', { method: 'POST' }); } catch(e) {}
      cargarEstadoCatalogo();
    }

    async function cargarEstadoCatalogo() {
      try {
        const res = await fetch('/api/catalogo/estado');
        let info = await res.json();
        if (info.necesitaRecarga && info.mapeoConfigurado) {
          await procesarExcel();
          const res2 = await fetch('/api/catalogo/estado');
          info = await res2.json();
        }
        renderCatalogoInfo(info);
        if (!info.mapeoConfigurado && document.getElementById('configOverlay').style.display !== 'flex') {
          abrirConfig();
        }
      } catch(e) {}
    }

    // Cobrado por el vendedor (desde el movil) y la caja aun no reviso el dinero.
    function necesitaRevision(p) {
      return p.estado === 'cobrado' && p.cobradoPor === 'vendedor' && p.revisado === false;
    }

    async function marcarRevisado(id, esperado) {
      const input = document.getElementById('revisar-monto-' + id);
      const entregado = input ? (parseFloat(input.value) || 0) : esperado;
      try {
        const res = await fetch('/api/pedidos/' + id + '/revisar', { method:'POST', body: JSON.stringify({ entregado }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          alert((data && data.error) || 'No se pudo marcar como revisado.');
        }
        cargarPedidos();
      } catch(e) { alert('No se pudo marcar como revisado.'); }
    }

    async function reportarProblema(id) {
      const motivo = prompt('¿Que no te entrego el vendedor (dinero, un producto, etc.)? Puedes dejarlo en blanco.', '');
      if (motivo === null) return;
      try {
        const res = await fetch('/api/pedidos/' + id + '/problema', { method:'POST', body: JSON.stringify({ motivo }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) { alert((data && data.error) || 'No se pudo enviar el aviso.'); return; }
        mostrarBanner('Se le aviso al vendedor sobre el problema con el pedido #' + id);
        cargarPedidos();
      } catch(e) { alert('No se pudo enviar el aviso (revisa la conexion).'); }
    }

    // ---- Fotos sin parpadeo: solo se pide /foto/<sku>.jpg si el producto TIENE foto ----
    let skusConFoto = null;              // Set con los SKU que tienen foto en la PC (null = aun no se sabe)
    const fotosFallidas = new Set();     // SKU cuya foto fallo una vez: no se vuelve a pedir
    try { const g = JSON.parse(localStorage.getItem('skusConFotoPC') || 'null'); if (Array.isArray(g)) skusConFoto = new Set(g.map(String)); } catch (e) {}
    async function cargarSkusConFoto() {
      try {
        const r = await fetch('/api/fotos/skus');
        const d = await r.json();
        if (Array.isArray(d)) {
          skusConFoto = new Set(d.map(String));
          try { localStorage.setItem('skusConFotoPC', JSON.stringify(d)); } catch (e) {}
        }
      } catch (e) {}
    }
    function tieneFoto(sku) {
      if (!sku) return false;
      sku = String(sku);
      if (fotosFallidas.has(sku)) return false;
      return skusConFoto ? skusConFoto.has(sku) : true;
    }
    function fotoFallo(img, sku) { fotosFallidas.add(String(sku)); if (img && img.remove) img.remove(); }
    // Cambia el contenido solo si de verdad es distinto: evita que la lista se "repinte" sola cada pocos segundos.
    function ponerHtmlSiCambio(el, html) {
      const hijo = el.firstElementChild;
      if (html !== '' && hijo && hijo.__marcaHtml === html && el.__htmlPrev === html) return false;
      el.innerHTML = html;
      el.__htmlPrev = html;
      if (el.firstElementChild) el.firstElementChild.__marcaHtml = html;
      return true;
    }
    cargarSkusConFoto();
    setInterval(cargarSkusConFoto, 60000);

    function renderCard(p) {
      const revisar = necesitaRevision(p);
      const estadoClass = revisar ? 'porrevisar' : (p.estado === 'cobrado' ? 'cobrado' : (p.estado === 'cancelado' ? 'cancelado' : 'pendiente'));
      const items = (p.items || []).map(it => {
        const fotoIt = tieneFoto(it.sku) ? ('<img class="ped-miniatura" src="/foto/' + encodeURIComponent(it.sku) + '.jpg" onerror=\'fotoFallo(this,' + JSON.stringify(String(it.sku)) + ')\' onclick=\'verFotoProductoNP(' + JSON.stringify(it.sku) + ',' + JSON.stringify(it.nombre || '') + ')\'>') : '';
        return '<li>' + fotoIt + '<span>' + it.cantidad + ' x ' + it.nombre + ' — $' + (it.precio * it.cantidad).toFixed(2) + '</span></li>';
      }).join('');

      // Si este pedido venia de un "para todos" que alguien tomo, se deja
      // constancia de quien fue y a que hora (por si dos vendedores discuten
      // quien iba a atender a un cliente).
      let origenHtml = '';
      if (Array.isArray(p.origenAsignados) && p.origenAsignados.length > 0) {
        origenHtml = p.origenAsignados.map(o => {
          const quien = o.tomadoPor || p.vendedor || 'un vendedor';
          const horaTom = o.horaTomado ? String(o.horaTomado).slice(11) : '';
          const clienteTxt = o.cliente ? (' — Cliente: <b>' + o.cliente + '</b>') : '';
          return '<div style="font-size:12px; color:#475569;">&#128274; Pedido "para todos" tomado por <b>' + quien + '</b>' + (horaTom ? (' a las ' + horaTom) : '') + clienteTxt + '.</div>';
        }).join('');
      }

      const notaHtml = p.nota ? ('<div style="font-size:13px; color:#0f172a; background:#e0f2fe; border-radius:6px; padding:6px 8px; margin:6px 0;">&#128172; <b>Nota del vendedor:</b> ' + escaparHtml(p.nota) + '</div>') : '';

      const totalProductos = Number(p.totalProductos !== undefined ? p.totalProductos : p.total || 0);
      const totalCobrado = Number(p.totalCobrado !== undefined && p.totalCobrado !== null ? p.totalCobrado : totalProductos);

      let totalHtml = '<div class="total">Total: $' + totalProductos.toFixed(2) + '</div>';
      if (Math.abs(totalCobrado - totalProductos) > 0.009) {
        totalHtml += '<div class="total" style="color:#7c2d12;">A cobrar (' + (p.metodoPago || '') + '): $' + totalCobrado.toFixed(2) + '</div>';
      }

      let extra = '';
      if (p.estado === 'cobrado' && p.metodoPago) {
        extra = '<div style="font-size:12px;">Pago: ' + p.metodoPago + (p.metodoPago === 'Combinado' ? (' (efectivo $' + Number(p.pagoEfectivo || 0).toFixed(2) + ' + transferencia $' + Number(p.pagoTransferencia || 0).toFixed(2) + ' x2)') : '') + '</div>';
      } else if (p.estado === 'pendiente' && p.metodoPago) {
        extra = '<div style="font-size:12px;">Metodo elegido por el vendedor: <b>' + p.metodoPago + '</b></div>';
      }

      let accionesPendiente = '';
      if (p.estado === 'pendiente') {
        accionesPendiente =
          '<button class="btn" onclick="cobrar(' + p.id + ', \'' + (p.metodoPago || 'Efectivo') + '\', ' + totalCobrado + ')">Cobrar en Caja (' + (p.metodoPago || 'Efectivo') + ') — $' + totalCobrado.toFixed(2) + '</button>' +
          (p.metodoPago === 'Combinado' ? '' : '<button class="btn" style="background:#0369a1;" onclick="abrirNuevoPedido(' + p.id + ')">+ Agregar producto</button>') +
          '<button class="btn btn-cancelar" onclick="cancelarPedido(' + p.id + ')">Cancelar pedido</button>';
      }

      // Pedido cobrado por el vendedor: la caja debe comprobar el dinero.
      let revisionHtml = '';
      let accionesRevision = '';
      if (revisar) {
        const metodo = p.metodoPago || 'Efectivo';
        const monto = '$' + totalCobrado.toFixed(2);
        let instruccion;
        if (metodo === 'Transferencia') instruccion = 'Verifica que llego la transferencia de ' + monto;
        else if (metodo === 'Efectivo') instruccion = 'El vendedor debe entregarte ' + monto + ' en efectivo';
        else instruccion = 'Verifica el pago (' + metodo + ') de ' + monto;
        let detalle = '';
        if (Number(p.montoRecibido) > 0) {
          detalle += '<div>Recibido del cliente: $' + Number(p.montoRecibido).toFixed(2) +
            (p.cambio !== null && p.cambio !== undefined && p.cambio !== '' ? ' — Cambio: $' + Number(p.cambio).toFixed(2) : '') + '</div>';
        }
        if (p.horaCobro) detalle += '<div>Cobrado por el vendedor a las ' + String(p.horaCobro).slice(11) + '</div>';
        let discrepanciaHtml = '';
        if (p.discrepancia) {
          discrepanciaHtml = '<div class="discrepancia-aviso">&#9888; El ultimo monto contado ($' +
            Number(p.montoEntregadoCaja || 0).toFixed(2) + ') no coincidia con lo que el vendedor tenia que entregar (' + monto +
            '). Ya se le aviso al vendedor. Vuelve a contar el dinero y confirma de nuevo.</div>';
        }
        let problemaHtml = '';
        if (p.problemaReportado) {
          problemaHtml = '<div class="discrepancia-aviso">&#9888; Reportaste un problema con este pedido' +
            (p.problemaMotivo ? (': ' + p.problemaMotivo) : '') + '. Ya se le aviso al vendedor.</div>';
        }
        revisionHtml = '<div class="revision"><b>' + instruccion + '</b>' + detalle + '</div>' + discrepanciaHtml + problemaHtml;
        accionesRevision =
          '<div class="revision-cuadre">' +
            '<input type="number" step="0.01" min="0" id="revisar-monto-' + p.id + '" value="' + totalCobrado.toFixed(2) + '" placeholder="Monto contado">' +
            '<button class="btn btn-ok" onclick="marcarRevisado(' + p.id + ', ' + totalCobrado + ')">Confirmar dinero</button>' +
            '<button class="btn btn-cancelar" onclick="reportarProblema(' + p.id + ')">No me entrego lo que era</button>' +
          '</div>';
      }

      const etiquetaEstado = revisar ? 'COBRADO POR EL VENDEDOR — REVISAR' : (p.estado === 'cobrado' ? 'COBRADO' : (p.estado === 'cancelado' ? 'CANCELADO' : 'PENDIENTE'));

      return '<div class="card ' + estadoClass + '">' +
        '<div class="estado">' + etiquetaEstado + '</div>' +
        '<h3>' + (p.vendedor || 'Vendedor') + '</h3>' +
        '<div class="hora">' + p.hora + ' — Folio #' + p.id + '</div>' +
        origenHtml +
        notaHtml +
        '<ul>' + items + '</ul>' +
        totalHtml +
        extra +
        revisionHtml +
        accionesRevision +
        accionesPendiente +
        (p.estado !== 'cancelado' ? '<button class="btn btn-imprimir" onclick="imprimir(' + p.id + ')">Imprimir recibo</button>' : '') +
        '</div>';
    }

    async function cancelarPedido(id) {
      if (!confirm('¿Cancelar este pedido? El stock de sus productos se devuelve automaticamente.')) return;
      try {
        const res = await fetch('/api/pedidos/' + id + '/cancelar', { method:'POST' });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) alert((data && data.error) || 'No se pudo cancelar el pedido.');
        cargarPedidos();
      } catch(e) { alert('No se pudo cancelar el pedido.'); }
    }

    async function cerrarDia() {
      if (!confirm('Esto archiva los pedidos cobrados y cancelados en la carpeta "historial" y limpia la lista. Los pendientes y los cobrados por vendedores que aun no marcaste como revisados se quedan. ¿Continuar?')) return;
      try {
        const res = await fetch('/api/dia/cerrar', { method:'POST' });
        const data = await res.json();
        if (data.ok) { alert('Dia cerrado. Quedan ' + data.restantes + ' pedido(s) pendiente(s) o por revisar.'); cargarPedidos(); }
        else { alert('No se pudo cerrar el dia: ' + (data.error || '')); }
      } catch(e) { alert('No se pudo cerrar el dia.'); }
    }

    async function cargarPedidos() {
      try {
        const res = await fetch('/api/pedidos');
        let pedidos = await res.json();
        pedidos.sort((a,b) => b.id - a.id);

        const soloPendientes = document.getElementById('soloPendientes').checked;
        const visibles = soloPendientes ? pedidos.filter(p => p.estado === 'pendiente' || necesitaRevision(p)) : pedidos;

        // Contador y aviso de pedidos cobrados por vendedores que la caja debe revisar
        const porRevisar = pedidos.filter(necesitaRevision);
        const contador = document.getElementById('contadorRevisar');
        contador.style.display = porRevisar.length ? 'inline-block' : 'none';
        contador.textContent = porRevisar.length + ' por revisar';
        const idsRevisar = new Set(porRevisar.map(p => p.id));
        const nuevoRevisar = primerCarga ? null : porRevisar.find(p => !idsRevisarAnt.has(p.id));
        idsRevisarAnt = idsRevisar;

        const grid = document.getElementById('grid');
        if (visibles.length === 0) {
          ponerHtmlSiCambio(grid, '<div class="vacio">No hay pedidos por mostrar.</div>');
        } else {
          ponerHtmlSiCambio(grid, visibles.map(renderCard).join(''));
        }

        const maxId = pedidos.reduce((m,p) => Math.max(m, p.id), 0);
        if (nuevoRevisar) {
          beep();
          mostrarBanner((nuevoRevisar.vendedor || 'El vendedor') + ' cobro el pedido #' + nuevoRevisar.id + ': revisa el dinero');
        } else if (!primerCarga && maxId > ultimoMaxId) {
          beep();
          mostrarBanner('Nuevo pedido recibido');
        }
        ultimoMaxId = maxId;
        primerCarga = false;
      } catch(e) {}
    }

    // ---- Ajuste: ocultar a los vendedores los productos sin stock ----
    let ocultarSinStock = false;
    let autoservicioDestino = 'pc';

    function pintarBotonSinStock() {
      const b = document.getElementById('btnOcultarSinStock');
      b.classList.toggle('activo', ocultarSinStock);
      b.textContent = 'Ocultar sin stock a vendedores: ' + (ocultarSinStock ? 'ACTIVADO' : 'DESACTIVADO');
    }

    function pintarBotonAutoservicio() {
      const b = document.getElementById('btnAutoservicioDestino');
      const aVendedor = autoservicioDestino === 'vendedor';
      b.classList.toggle('activo', aVendedor);
      b.textContent = 'Pedidos de clientes (autoservicio): ' + (aVendedor ? 'A cualquier vendedor disponible' : 'Directo a la PC');
    }

    let permitirDescuentos = false;
    function pintarBotonDescuentos() {
      const b = document.getElementById('btnPermitirDescuentos');
      b.classList.toggle('activo', permitirDescuentos);
      b.textContent = 'Descuentos por producto: ' + (permitirDescuentos ? 'ACTIVADOS' : 'DESACTIVADOS');
    }
    async function alternarPermitirDescuentos() {
      try {
        const res = await fetch('/api/config', { method:'POST', body: JSON.stringify({ permitirDescuentos: !permitirDescuentos }) });
        const cfg = await res.json();
        permitirDescuentos = !!cfg.permitirDescuentos;
        pintarBotonDescuentos();
      } catch (e) { alert('No se pudo guardar el ajuste.'); }
    }

    async function cargarConfigApp() {
      try {
        const res = await fetch('/api/config');
        const cfg = await res.json();
        permitirDescuentos = !!cfg.permitirDescuentos;
        pintarBotonDescuentos();
        ocultarSinStock = !!cfg.ocultarSinStock;
        pintarBotonSinStock();
        autoservicioDestino = cfg.autoservicioDestino === 'vendedor' ? 'vendedor' : 'pc';
        pintarBotonAutoservicio();
        const inTasa = document.getElementById('tasaDolar');
        if (document.activeElement !== inTasa) inTasa.value = cfg.tasaDolar ? cfg.tasaDolar : '';
        tasaDolarPC = parseFloat(cfg.tasaDolar) || 0;
        const inUmbral = document.getElementById('umbralStockBajo');
        if (document.activeElement !== inUmbral) inUmbral.value = (cfg.umbralStockBajo !== undefined && cfg.umbralStockBajo !== null) ? cfg.umbralStockBajo : 3;
        actualizarAlertaMenuPC();
      } catch(e) {}
    }

    async function guardarUmbralStockBajo() {
      const v = parseInt(document.getElementById('umbralStockBajo').value, 10);
      try {
        await fetch('/api/config', { method:'POST', body: JSON.stringify({ umbralStockBajo: isNaN(v) ? 3 : v }) });
        revisarStockBajo();
      } catch (e) { alert('No se pudo guardar el ajuste.'); }
    }

    // Punto rojo en "Ajustes": avisa sin tener que abrir el menu si el
    // catalogo tiene error o si la Tasa USD quedo en 0 (no se imprime en recibos).
    let catalogoConError = false;
    let tasaDolarPC = 0;
    function actualizarAlertaMenuPC() {
      const btn = document.getElementById('btnMenuPC');
      if (!btn) return;
      btn.classList.toggle('alerta', catalogoConError || tasaDolarPC <= 0);
    }

    async function guardarTasaDolar() {
      const v = parseFloat(String(document.getElementById('tasaDolar').value).replace(',', '.')) || 0;
      try {
        const res = await fetch('/api/config', { method:'POST', body: JSON.stringify({ tasaDolar: v }) });
        if (!res.ok) throw new Error('http ' + res.status);
        mostrarBanner(v > 0 ? ('Tasa USD guardada: ' + v) : 'Tasa USD desactivada: no se imprime TOTAL USD');
      } catch(e) { alert('No se pudo guardar la tasa.'); }
    }

    async function alternarOcultarSinStock() {
      const nuevo = !ocultarSinStock;
      try {
        const res = await fetch('/api/config', { method:'POST', body: JSON.stringify({ ocultarSinStock: nuevo }) });
        if (!res.ok) throw new Error('http ' + res.status);
        const cfg = await res.json();
        ocultarSinStock = !!cfg.ocultarSinStock;
        pintarBotonSinStock();
        mostrarBanner(ocultarSinStock
          ? 'Activado: los vendedores ya no ven productos sin stock'
          : 'Desactivado: los vendedores ven todos los productos');
      } catch(e) { alert('No se pudo guardar el ajuste.'); }
    }

    async function alternarAutoservicioDestino() {
      const nuevo = autoservicioDestino === 'vendedor' ? 'pc' : 'vendedor';
      try {
        const res = await fetch('/api/config', { method:'POST', body: JSON.stringify({ autoservicioDestino: nuevo }) });
        if (!res.ok) throw new Error('http ' + res.status);
        const cfg = await res.json();
        autoservicioDestino = cfg.autoservicioDestino === 'vendedor' ? 'vendedor' : 'pc';
        pintarBotonAutoservicio();
        mostrarBanner(autoservicioDestino === 'vendedor'
          ? 'Los pedidos de clientes ahora se ofrecen a cualquier vendedor conectado'
          : 'Los pedidos de clientes ahora caen directo a este panel');
      } catch(e) { alert('No se pudo guardar el ajuste.'); }
    }

    let enlaceClienteCache = null;
    async function obtenerEnlaceCliente() {
      if (enlaceClienteCache) return enlaceClienteCache;
      let base = location.origin;
      try {
        const res = await fetch('/api/ip');
        const info = await res.json();
        if (info && info.ip) base = 'http://' + info.ip + ':' + info.puerto;
      } catch (e) { /* si falla, se usa location.origin (sirve solo en esta PC) */ }
      enlaceClienteCache = base + '/vendedor?cliente=1';
      return enlaceClienteCache;
    }

    async function copiarEnlaceAutoservicio() {
      const enlace = await obtenerEnlaceCliente();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(enlace)
          .then(() => mostrarBanner('Enlace copiado: ' + enlace))
          .catch(() => prompt('Copia este enlace para compartirlo con los clientes:', enlace));
      } else {
        prompt('Copia este enlace para compartirlo con los clientes:', enlace);
      }
    }

    function dibujarQREnBox(texto, box) {
      box.innerHTML = '';
      try {
        const qr = qrcode(0, 'M');
        qr.addData(texto);
        qr.make();
        const canvas = document.createElement('canvas');
        const tam = 5;
        const margen = tam * 4;
        const lado = qr.getModuleCount() * tam + margen * 2;
        canvas.width = lado; canvas.height = lado;
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, lado, lado);
        ctx.save(); ctx.translate(margen, margen);
        qr.renderTo2dContext(ctx, tam);
        ctx.restore();
        box.appendChild(canvas);
        return true;
      } catch (e) {
        box.textContent = 'No se pudo generar el QR.';
        return false;
      }
    }

    let enlaceQRActual = null;
    async function mostrarQRCliente() {
      const enlace = await obtenerEnlaceCliente();
      enlaceQRActual = enlace;
      dibujarQREnBox(enlace, document.getElementById('qrCodeBox'));
      document.getElementById('qrTituloModal').textContent = 'QR para clientes';
      document.getElementById('qrEnlaceTexto').textContent = enlace;
      document.getElementById('qrBotonCopiar').style.display = 'block';
      document.getElementById('qrOverlay').style.display = 'flex';
    }

    let enlaceVendedorCache = null;
    async function obtenerEnlaceVendedor() {
      if (enlaceVendedorCache) return enlaceVendedorCache;
      let base = location.origin;
      try {
        const res = await fetch('/api/ip');
        const info = await res.json();
        if (info && info.ip) base = 'http://' + info.ip + ':' + info.puerto;
      } catch (e) { /* si falla, se usa location.origin (sirve solo en esta PC) */ }
      enlaceVendedorCache = base + '/vendedor';
      return enlaceVendedorCache;
    }
    async function mostrarQRVendedor() {
      const enlace = await obtenerEnlaceVendedor();
      enlaceQRActual = enlace;
      dibujarQREnBox(enlace, document.getElementById('qrCodeBox'));
      document.getElementById('qrTituloModal').textContent = 'QR para vendedores';
      document.getElementById('qrEnlaceTexto').textContent = enlace;
      document.getElementById('qrBotonCopiar').style.display = 'block';
      document.getElementById('qrOverlay').style.display = 'flex';
    }
    function copiarEnlaceQRActual() {
      const enlace = enlaceQRActual;
      if (!enlace) return;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(enlace).then(() => mostrarBanner('Enlace copiado: ' + enlace)).catch(() => prompt('Copia este enlace:', enlace));
      } else {
        prompt('Copia este enlace:', enlace);
      }
    }

    function escaparWifi(v) {
      return String(v).replace(/([\\;,":])/g, '\\$1');
    }

    function construirTextoWifi(ssid, clave) {
      const seguridad = clave ? 'WPA' : 'nopass';
      return 'WIFI:T:' + seguridad + ';S:' + escaparWifi(ssid) + ';' + (clave ? ('P:' + escaparWifi(clave) + ';') : '') + ';';
    }

    async function guardarConfigWifi(ssid, clave) {
      try {
        await fetch('/api/config', { method: 'POST', body: JSON.stringify({ wifiSSID: ssid, wifiClave: clave }) });
      } catch (e) { /* si falla el guardado, se sigue usando lo escrito en pantalla */ }
    }

    async function cargarDatosWifi() {
      try {
        const res = await fetch('/api/config');
        const cfg = await res.json();
        document.getElementById('inputWifiSSID').value = cfg.wifiSSID || '';
        document.getElementById('inputWifiClave').value = cfg.wifiClave || '';
      } catch (e) { /* si falla, se dejan los campos vacios */ }
    }

    async function guardarYMostrarQRWifi() {
      const ssid = document.getElementById('inputWifiSSID').value.trim();
      const clave = document.getElementById('inputWifiClave').value;
      if (!ssid) { alert('Escribe el nombre (SSID) de la red WiFi.'); return; }
      await guardarConfigWifi(ssid, clave);
      dibujarQREnBox(construirTextoWifi(ssid, clave), document.getElementById('qrCodeBox'));
      document.getElementById('qrTituloModal').textContent = 'QR para conectarse al WiFi';
      document.getElementById('qrEnlaceTexto').textContent = 'Red: ' + ssid;
      document.getElementById('qrBotonCopiar').style.display = 'none';
      document.getElementById('qrOverlay').style.display = 'flex';
    }

    async function generarTarjetaImprimir() {
      const ssid = document.getElementById('inputWifiSSID').value.trim();
      const clave = document.getElementById('inputWifiClave').value;
      if (!ssid) { alert('Escribe el nombre (SSID) de la red WiFi antes de generar la tarjeta.'); return; }
      await guardarConfigWifi(ssid, clave);
      const enlace = await obtenerEnlaceCliente();
      dibujarQREnBox(construirTextoWifi(ssid, clave), document.getElementById('tarjetaQRWifi'));
      dibujarQREnBox(enlace, document.getElementById('tarjetaQRCatalogo'));
      document.getElementById('tarjetaWifiNombre').textContent = 'Red: ' + ssid;
      setTimeout(() => window.print(), 150);
    }

    function cerrarQRCliente() {
      document.getElementById('qrOverlay').style.display = 'none';
    }

    // ---- Fotos del catalogo (desde el backup de la app del catalogo) ----
    async function pintarBotonFotos() {
      try {
        const res = await fetch('/api/fotos/estado');
        const info = await res.json();
        document.getElementById('btnFotosCatalogo').textContent = 'Fotos del catalogo: ' + info.cantidadFotos;
      } catch(e) {
        document.getElementById('btnFotosCatalogo').textContent = 'Fotos del catalogo: ...';
      }
    }

    async function importarFotosCatalogo() {
      const b = document.getElementById('btnFotosCatalogo');
      const original = b.textContent;
      b.textContent = 'Buscando backup...';
      b.disabled = true;
      try {
        const res = await fetch('/api/fotos/importar', { method: 'POST' });
        const info = await res.json();
        if (!res.ok || !info.ok) {
          alert(info.error || 'No se pudo importar las fotos.');
        } else {
          mostrarBanner('Fotos actualizadas desde ' + info.archivo + ': ' + info.copiadas + ' de ' + info.totalProductos + ' productos' + (info.sinFoto ? (', ' + info.sinFoto + ' sin foto en el catalogo') : '') + (info.conError ? (', ' + info.conError + ' con error') : ''));
        }
      } catch(e) {
        alert('No se pudo importar las fotos.');
      } finally {
        b.disabled = false;
        pintarBotonFotos();
      }
    }

    document.getElementById('zipFotosPC').addEventListener('change', async (ev) => {
      const file = ev.target.files[0];
      ev.target.value = '';
      if (!file) return;
      const b = document.getElementById('btnFotosCatalogo');
      const original = b.textContent;
      b.textContent = 'Subiendo y procesando...';
      b.disabled = true;
      try {
        const buffer = await file.arrayBuffer();
        const res = await fetch('/api/fotos/subir', { method: 'POST', body: buffer, headers: { 'Content-Type': 'application/octet-stream' } });
        const info = await res.json();
        if (!res.ok || !info.ok) {
          alert(info.error || 'No se pudo procesar el archivo.');
        } else {
          mostrarBanner('Fotos actualizadas desde ' + file.name + ': ' + info.copiadas + ' de ' + info.totalProductos + ' productos' + (info.sinFoto ? (', ' + info.sinFoto + ' sin foto en el catalogo') : '') + (info.conError ? (', ' + info.conError + ' con error') : ''));
        }
      } catch (e) {
        alert('No se pudo subir el archivo.');
      } finally {
        b.disabled = false;
        pintarBotonFotos();
      }
    });

    document.getElementById('soloPendientes').addEventListener('change', cargarPedidos);
    cargarPedidos();
    cargarEstadoCatalogo();
    cargarConfigApp();
    pintarBotonFotos();
    cargarDatosWifi();
    setInterval(cargarPedidos, 2000);

    // ---- Alerta de stock bajo: campanita con lista. Cada aviso se puede borrar;
    // un producto borrado vuelve a aparecer solo si su stock cambia (baja mas).
    // Suena/avisa solo cuando aparece un SKU nuevo por debajo del umbral.
    let stockBajoLista = [];
    let skusStockBajoAnt = new Set();
    let primerSondeoStockBajo = true;
    function descartadosStockBajo() { try { return JSON.parse(localStorage.getItem('stockBajoDescartados') || '{}'); } catch (e) { return {}; } }
    function guardarDescartadosStockBajo(o) { try { localStorage.setItem('stockBajoDescartados', JSON.stringify(o)); } catch (e) {} }
    function visiblesStockBajo() {
      const d = descartadosStockBajo();
      return stockBajoLista.filter(p => !(p.sku in d) || Number(d[p.sku]) !== Number(p.stock));
    }
    function pintarCampanaStockBajo() {
      const v = visiblesStockBajo();
      const badge = document.getElementById('badgeStockBajo');
      badge.style.display = v.length ? 'flex' : 'none';
      badge.textContent = v.length;
      document.getElementById('listaStockBajo').innerHTML = v.length
        ? v.map(p => '<div style="display:flex; justify-content:space-between; align-items:center; gap:8px; padding:6px 0; border-bottom:1px solid #334155;"><span>&#9888; ' + (p.nombre || p.sku) + ' <b>(' + p.stock + ')</b></span><button class="btn-toggle" style="font-size:12px; padding:3px 8px;" onclick=\'descartarStockBajo(' + JSON.stringify(p.sku) + ')\'>&#10005;</button></div>').join('')
        : '<div style="color:#94a3b8;">Sin avisos de stock bajo.</div>';
    }
    function togglePanelStockBajo() {
      const el = document.getElementById('panelStockBajo');
      el.style.display = el.style.display === 'none' ? 'block' : 'none';
    }
    function descartarStockBajo(sku) {
      const p = stockBajoLista.find(x => x.sku === sku);
      if (!p) return;
      const d = descartadosStockBajo(); d[sku] = p.stock; guardarDescartadosStockBajo(d);
      pintarCampanaStockBajo();
    }
    function descartarTodoStockBajo() {
      const d = descartadosStockBajo();
      stockBajoLista.forEach(p => { d[p.sku] = p.stock; });
      guardarDescartadosStockBajo(d);
      pintarCampanaStockBajo();
    }
    async function revisarStockBajo() {
      try {
        const res = await fetch('/api/stockbajo');
        const data = await res.json();
        stockBajoLista = (data && data.umbral) ? ((data && data.productos) || []) : [];
        const skusAhora = new Set(stockBajoLista.map(p => p.sku));
        const d = descartadosStockBajo();
        const nuevos = stockBajoLista.filter(p => !skusStockBajoAnt.has(p.sku) && (!(p.sku in d) || Number(d[p.sku]) !== Number(p.stock)));
        if (!primerSondeoStockBajo && nuevos.length > 0) {
          beep();
          mostrarBanner('Stock bajo: ' + nuevos.map(p => p.nombre || p.sku).join(', '));
        }
        skusStockBajoAnt = skusAhora;
        primerSondeoStockBajo = false;
        pintarCampanaStockBajo();
      } catch (e) {}
    }
    revisarStockBajo();
    setInterval(revisarStockBajo, 15000);
    setInterval(cargarEstadoCatalogo, 15000);
    setInterval(cargarConfigApp, 15000);

    // Quita acentos y pasa a minusculas: "bateria" encuentra "Batería" y viceversa.
    function normalizar(t) {
      return String(t == null ? '' : t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    }

    // ---- PIN de vendedores ----
    // =====================================================================
    // Lista de compras / reabastecimiento
    // =====================================================================
    let comprasItems = [];
    let catalogoCompras = [];
    function pcById(id) { return document.getElementById(id); }

    async function abrirCompras() {
      pcById('menuPCOverlay').style.display = 'none';
      pcById('comprasOverlay').style.display = 'flex';
      pcById('comprasMsg').textContent = '';
      pcById('minBuscador').value = '';
      pcById('minResultados').innerHTML = '';
      try { const r = await fetch('/api/catalogo'); catalogoCompras = await r.json(); } catch (e) { catalogoCompras = []; }
      await cargarCompras();
    }
    function cerrarCompras() { pcById('comprasOverlay').style.display = 'none'; }

    async function cargarCompras() {
      pcById('comprasLista').innerHTML = '<div style="color:#94a3b8; font-size:13px; padding:8px;">Cargando...</div>';
      try {
        const r = await fetch('/api/reabastecer');
        const d = await r.json();
        comprasItems = (d.items || []).map(i => Object.assign({}, i, { comprar: i.sugerido, incluir: true }));
      } catch (e) { comprasItems = []; }
      renderCompras();
    }

    function renderCompras() {
      const c = pcById('comprasLista');
      if (comprasItems.length === 0) {
        c.innerHTML = '<div style="color:#94a3b8; font-size:13px; padding:10px;">Nada por reabastecer: ningún producto está en su mínimo.</div>';
      } else {
        c.innerHTML = comprasItems.map((i, idx) =>
          '<div class="np-resultado"><label style="display:flex; gap:8px; align-items:center; flex:1; min-width:0;">' +
          '<input type="checkbox" ' + (i.incluir ? 'checked' : '') + ' onchange="comprasItems[' + idx + '].incluir = this.checked; resumenCompras()">' +
          '<span style="min-width:0;">' + escaparHtml(i.nombre) + '<br><small style="color:#94a3b8;">' + escaparHtml(i.sku || 'sin SKU') + ' · hay ' + i.stock + ' · mín. ' + i.minimo + (i.personalizado ? ' (propio)' : '') + '</small></span></label>' +
          '<input type="number" min="1" value="' + i.comprar + '" style="width:72px; padding:6px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0;" onchange="comprasItems[' + idx + '].comprar = Math.max(1, parseFloat(this.value) || 1); resumenCompras()">' +
          '</div>'
        ).join('');
      }
      resumenCompras();
    }

    function comprasSeleccionadas() { return comprasItems.filter(i => i.incluir && i.comprar > 0); }
    function resumenCompras() {
      const s = comprasSeleccionadas();
      let u = 0; s.forEach(i => { u += Number(i.comprar) || 0; });
      pcById('comprasResumen').textContent = s.length + ' artículo(s) en la lista · ' + u + ' unidades a comprar';
    }
    function textoCompras() {
      return 'LISTA DE COMPRAS ' + new Date().toLocaleDateString('es-ES') + '\n' +
        comprasSeleccionadas().map(i => '- ' + i.comprar + ' x ' + i.nombre + (i.sku ? (' [' + i.sku + ']') : '')).join('\n');
    }
    async function copiarCompras() {
      const t = textoCompras();
      try { await navigator.clipboard.writeText(t); }
      catch (e) {
        const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); } catch (e2) {}
        ta.remove();
      }
      pcById('comprasMsg').style.color = '#86efac';
      pcById('comprasMsg').textContent = 'Lista copiada: pégala en WhatsApp o donde quieras.';
    }
    async function imprimirCompras() {
      const s = comprasSeleccionadas();
      const msg = pcById('comprasMsg');
      if (s.length === 0) { msg.style.color = '#fca5a5'; msg.textContent = 'No hay artículos marcados.'; return; }
      try {
        const r = await fetch('/api/reabastecer/imprimir', { method: 'POST', body: JSON.stringify({
          items: s.map(i => ({ sku: i.sku, nombre: i.nombre, cantidad: i.comprar, stock: i.stock, minimo: i.minimo }))
        }) });
        const d = await r.json();
        msg.style.color = d.ok ? '#86efac' : '#fca5a5';
        msg.textContent = d.ok ? 'Lista enviada a la impresora.' : (d.error || 'No se pudo imprimir.');
      } catch (e) { msg.style.color = '#fca5a5'; msg.textContent = 'No se pudo contactar con el servidor.'; }
    }
    function hojaCompras() {
      const s = comprasSeleccionadas();
      if (s.length === 0) { pcById('comprasMsg').style.color = '#fca5a5'; pcById('comprasMsg').textContent = 'No hay artículos marcados.'; return; }
      const w = window.open('', '_blank');
      if (!w) return;
      w.document.write('<html><head><meta charset="UTF-8"><title>Lista de compras</title><style>body{font-family:Arial,sans-serif;padding:20px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #999;padding:6px 8px;text-align:left;font-size:14px}th{background:#eee}</style></head><body><h2>Lista de compras - ' +
        new Date().toLocaleDateString('es-ES') + '</h2><table><tr><th>✓</th><th>Cant.</th><th>Producto</th><th>SKU</th><th>Hay</th><th>Mín.</th></tr>' +
        s.map(i => '<tr><td>&#9744;</td><td>' + i.comprar + '</td><td>' + escaparHtml(i.nombre) + '</td><td>' + escaparHtml(i.sku || '') + '</td><td>' + i.stock + '</td><td>' + i.minimo + '</td></tr>').join('') +
        '</table><script>window.onload=function(){window.print()}<\/script></body></html>');
      w.document.close();
    }

    function buscarMinimo() {
      const q = normalizar(pcById('minBuscador').value);
      const cont = pcById('minResultados');
      if (q.length < 2) { cont.innerHTML = ''; return; }
      const res = catalogoCompras.filter(p => p.sku && (normalizar(p.nombre).includes(q) || normalizar(p.sku).includes(q))).slice(0, 6);
      cont.innerHTML = res.length ? res.map((p, idx) =>
        '<div class="np-resultado"><span style="flex:1; min-width:0;">' + escaparHtml(p.nombre) + '<br><small style="color:#94a3b8;">' + escaparHtml(p.sku) + ' · hay ' + (p.stock == null ? '—' : p.stock) + '</small></span>' +
        '<input type="number" min="0" id="minIn' + idx + '" placeholder="mín." style="width:64px; padding:6px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; margin:0 6px;">' +
        '<button class="btn-toggle" style="padding:6px 10px; font-size:12px;" onclick="fijarMinimo(' + JSON.stringify(p.sku).replace(/"/g, '&quot;') + ', ' + idx + ')">Fijar</button>' +
        '<button class="btn-toggle" style="padding:6px 10px; font-size:12px; margin-left:4px;" title="Volver al mínimo general" onclick="fijarMinimo(' + JSON.stringify(p.sku).replace(/"/g, '&quot;') + ', -1)">Quitar</button></div>'
      ).join('') : '<div class="np-vacio">Sin resultados.</div>';
    }
    async function fijarMinimo(sku, idx) {
      let minimo = null;
      if (idx >= 0) { const v = pcById('minIn' + idx).value; if (v === '') { return; } minimo = parseFloat(v); }
      await fetch('/api/minimos', { method: 'POST', body: JSON.stringify({ sku, minimo }) });
      pcById('comprasMsg').style.color = '#86efac';
      pcById('comprasMsg').textContent = idx >= 0 ? 'Mínimo guardado.' : 'Ese producto vuelve a usar el mínimo general.';
      await cargarCompras();
      try { revisarStockBajo(); } catch (e) {}
    }

    // =====================================================================
    // Devolucion / cambio / garantia
    // =====================================================================
    let devItems = [];
    let catalogoDev = [];
    let devResultadosLista = [];
    let devAtendio = '';

    async function abrirDevolucion() {
      pcById('menuPCOverlay').style.display = 'none';
      pcById('devOverlay').style.display = 'flex';
      devItems = []; devAtendio = '';
      ['devPedido', 'devCliente', 'devMotivo', 'devEntregado', 'devDif', 'devBuscar'].forEach(id => { pcById(id).value = ''; });
      pcById('devResultados').innerHTML = '';
      pcById('devMsg').textContent = '';
      pcById('devTipo').value = 'devolucion';
      pcById('devStock').checked = true;
      try { const r = await fetch('/api/catalogo'); catalogoDev = await r.json(); } catch (e) { catalogoDev = []; }
      renderDev();
      cargarHistorialDev();
    }
    function cerrarDevolucion() { pcById('devOverlay').style.display = 'none'; }
    function devTipoCambio() { pcById('devStock').checked = (pcById('devTipo').value === 'devolucion'); }
    function devMensaje(t, ok) { const m = pcById('devMsg'); m.style.color = ok ? '#86efac' : '#fca5a5'; m.textContent = t; }

    function devBuscarProd() {
      const q = normalizar(pcById('devBuscar').value);
      const cont = pcById('devResultados');
      if (q.length < 2) { cont.innerHTML = ''; devResultadosLista = []; return; }
      devResultadosLista = catalogoDev.filter(p => normalizar(p.nombre).includes(q) || normalizar(p.sku || '').includes(q)).slice(0, 6);
      cont.innerHTML = devResultadosLista.map((p, idx) =>
        '<div class="np-resultado"><span style="flex:1; min-width:0;">' + escaparHtml(p.nombre) + '<br><small style="color:#94a3b8;">' + escaparHtml(p.sku || 'sin SKU') + ' · $' + Number(p.precio).toFixed(2) + '</small></span>' +
        '<button class="btn-toggle" style="padding:6px 12px;" onclick="devAgregar(' + idx + ')">+</button></div>'
      ).join('') || '<div class="np-vacio">Sin resultados.</div>';
    }
    function devAgregar(idx) {
      const p = devResultadosLista[idx];
      if (!p) return;
      const ya = devItems.find(i => i.sku && i.sku === p.sku);
      if (ya) ya.cantidad += 1; else devItems.push({ sku: p.sku || '', nombre: p.nombre, cantidad: 1, precio: Number(p.precio) || 0 });
      pcById('devBuscar').value = ''; pcById('devResultados').innerHTML = '';
      renderDev();
    }
    async function devCargarPedido() {
      const id = parseInt(pcById('devPedido').value, 10);
      if (!id) { devMensaje('Escribe el número del pedido.', false); return; }
      try {
        const r = await fetch('/api/pedidos');
        const lista = await r.json();
        const p = (Array.isArray(lista) ? lista : []).find(x => Number(x.id) === id);
        if (!p) { devMensaje('Ese pedido no está en los de hoy (si ya se archivó, agrega los productos a mano).', false); return; }
        const f = p.metodoPago === 'Transferencia' ? 2 : 1;
        devItems = (p.items || []).map(it => ({ sku: it.sku || '', nombre: it.nombre, cantidad: Number(it.cantidad) || 1, precio: (Number(it.precio) || 0) * f }));
        devAtendio = p.vendedor || '';
        if (p.metodoPago === 'Efectivo' || p.metodoPago === 'Transferencia') pcById('devMetodo').value = p.metodoPago;
        devMensaje('Pedido #' + id + ' cargado. Quita o ajusta las líneas que no se devuelven.', true);
        renderDev();
      } catch (e) { devMensaje('No se pudo leer el pedido.', false); }
    }
    function devTotalNum() { let t = 0; devItems.forEach(i => { t += (Number(i.precio) || 0) * (Number(i.cantidad) || 0); }); return t; }
    function renderDev() {
      const c = pcById('devLista');
      c.innerHTML = devItems.length ? devItems.map((i, idx) =>
        '<div class="np-resultado"><span style="flex:1; min-width:0;">' + escaparHtml(i.nombre) + '</span>' +
        '<input type="number" min="1" step="any" value="' + i.cantidad + '" title="Cantidad" style="width:58px; padding:6px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; margin:0 4px;" onchange="devItems[' + idx + '].cantidad = Math.max(0.01, parseFloat(this.value) || 1); renderDev()">' +
        '<input type="number" min="0" step="any" value="' + Number(i.precio).toFixed(2) + '" title="Precio" style="width:76px; padding:6px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0;" onchange="devItems[' + idx + '].precio = Math.max(0, parseFloat(this.value) || 0); renderDev()">' +
        '<button class="btn-toggle" style="padding:6px 10px; margin-left:4px;" onclick="devItems.splice(' + idx + ', 1); renderDev()">&times;</button></div>'
      ).join('') : '<div class="np-vacio">Sin productos aún.</div>';
      pcById('devTotal').textContent = devTotalNum().toFixed(2);
    }
    async function devGuardar() {
      if (devItems.length === 0) { devMensaje('Agrega al menos un producto.', false); return; }
      const dif = pcById('devDif').value;
      const body = {
        tipo: pcById('devTipo').value,
        pedidoId: pcById('devPedido').value.trim(),
        cliente: pcById('devCliente').value.trim(),
        atendio: devAtendio || 'Caja',
        motivo: pcById('devMotivo').value.trim(),
        items: devItems,
        metodoReembolso: pcById('devMetodo').value,
        entregado: pcById('devEntregado').value.trim(),
        diferencia: dif === '' ? null : parseFloat(dif),
        reintegrarStock: pcById('devStock').checked
      };
      try {
        const r = await fetch('/api/devoluciones', { method: 'POST', body: JSON.stringify(body) });
        const d = await r.json();
        if (!d.ok) { devMensaje(d.error || 'No se pudo guardar.', false); return; }
        devMensaje(d.impreso ? ('Comprobante D-' + d.id + ' impreso.') : ('Guardado como D-' + d.id + ', pero no se pudo imprimir: ' + (d.errorImpresion || 'revisa la impresora') + '. Usa Reimprimir abajo.'), d.impreso);
        devItems = []; devAtendio = '';
        ['devPedido', 'devCliente', 'devMotivo', 'devEntregado', 'devDif'].forEach(id => { pcById(id).value = ''; });
        renderDev(); cargarHistorialDev();
        try { cargarCatalogoNP(); } catch (e) {}
      } catch (e) { devMensaje('No se pudo contactar con el servidor.', false); }
    }
    async function cargarHistorialDev() {
      const cont = pcById('devHistorial');
      try {
        const r = await fetch('/api/devoluciones');
        const d = await r.json();
        const lista = (d.devoluciones || []).slice(0, 8);
        const nombres = { devolucion: 'Devolución', cambio: 'Cambio', garantia: 'Garantía' };
        cont.innerHTML = lista.length ? lista.map(x =>
          '<div class="np-resultado"><span style="flex:1; min-width:0;">D-' + x.id + ' · ' + (nombres[x.tipo] || x.tipo) + ' · $' + Number(x.total || 0).toFixed(2) + '<br><small style="color:#94a3b8;">' + escaparHtml(x.hora) + (x.cliente ? ' · ' + escaparHtml(x.cliente) : '') + '</small></span>' +
          '<button class="btn-toggle" style="padding:6px 10px; font-size:12px;" onclick="devReimprimir(' + x.id + ')">Reimprimir</button></div>'
        ).join('') : '<div class="np-vacio">Todavía no hay comprobantes.</div>';
      } catch (e) { cont.innerHTML = ''; }
    }
    async function devReimprimir(id) {
      try {
        const r = await fetch('/api/devoluciones/' + id + '/imprimir', { method: 'POST' });
        const d = await r.json();
        devMensaje(d.ok ? ('Comprobante D-' + id + ' enviado a la impresora.') : (d.error || 'No se pudo imprimir.'), !!d.ok);
      } catch (e) { devMensaje('No se pudo contactar con el servidor.', false); }
    }

    // =====================================================================
    // Permisos por vendedor
    // =====================================================================
    let permData = null;
    let permActual = '';
    const PERM_PRESETS = {
      completo: null,
      vender: { crearPedidos: true, cobrar: false, cancelar: false, editar: false, imprimir: true, descuentos: false, verStock: true, misPedidos: true, ajustes: false, modoCliente: true, asignados: true, notificaciones: true },
      consulta: { crearPedidos: false, cobrar: false, cancelar: false, editar: false, imprimir: false, descuentos: false, verStock: true, misPedidos: false, ajustes: false, modoCliente: true, asignados: false, notificaciones: true }
    };
    async function abrirPermisos() {
      pcById('menuPCOverlay').style.display = 'none';
      pcById('permOverlay').style.display = 'flex';
      pcById('permMsg').textContent = '';
      await cargarPermisos();
    }
    function cerrarPermisos() { pcById('permOverlay').style.display = 'none'; }
    async function cargarPermisos() {
      try { const r = await fetch('/api/permisos/todos'); permData = await r.json(); } catch (e) { permData = null; }
      const sel = pcById('permVendedor');
      if (!permData || !permData.vendedores || permData.vendedores.length === 0) {
        sel.innerHTML = '';
        pcById('permLista').innerHTML = '<div class="np-vacio">Todavía no hay vendedores. Créalos primero en "PIN de vendedores".</div>';
        return;
      }
      if (!permData.vendedores.some(v => v.nombre === permActual)) permActual = permData.vendedores[0].nombre;
      sel.innerHTML = permData.vendedores.map(v => '<option value="' + escaparHtml(v.nombre) + '"' + (v.nombre === permActual ? ' selected' : '') + '>' + escaparHtml(v.nombre) + '</option>').join('');
      renderPermisos();
    }
    function permSeleccionar(n) { permActual = n; renderPermisos(); }
    function renderPermisos() {
      const v = (permData.vendedores || []).find(x => x.nombre === permActual);
      if (!v) return;
      pcById('permLista').innerHTML = permData.claves.map(c =>
        '<label style="display:flex; align-items:flex-start; gap:10px; padding:8px 0; border-bottom:1px solid #334155; font-size:14px; cursor:pointer;">' +
        '<input type="checkbox" style="margin-top:3px;" ' + (v.permisos[c.k] ? 'checked' : '') + ' onchange="permCambiar(\'' + c.k + '\', this.checked)">' +
        '<span><b style="color:#fff;">' + escaparHtml(c.t) + '</b><br><small style="color:#94a3b8;">' + escaparHtml(c.d) + '</small></span></label>'
      ).join('');
    }
    async function permGuardar(permisos) {
      const m = pcById('permMsg');
      try {
        const r = await fetch('/api/permisos', { method: 'POST', body: JSON.stringify({ vendedor: permActual, permisos }) });
        const d = await r.json();
        m.style.color = d.ok ? '#86efac' : '#fca5a5';
        m.textContent = d.ok ? ('Guardado para ' + permActual + '.') : (d.error || 'No se pudo guardar.');
        if (d.ok) { const v = permData.vendedores.find(x => x.nombre === permActual); if (v) v.permisos = d.permisos; }
      } catch (e) { m.style.color = '#fca5a5'; m.textContent = 'No se pudo contactar con el servidor.'; }
      renderPermisos();
    }
    function permCambiar(k, valor) { const o = {}; o[k] = valor; permGuardar(o); }
    function permPreset(nombre) {
      const p = {};
      permData.claves.forEach(c => { p[c.k] = PERM_PRESETS[nombre] ? !!PERM_PRESETS[nombre][c.k] : true; });
      permGuardar(p);
    }

    // ---- Red / IP fija ----
    async function abrirRed() {
      pcById('menuPCOverlay').style.display = 'none';
      pcById('redOverlay').style.display = 'flex';
      pcById('redMsg').textContent = '';
      await cargarRed();
    }
    async function cargarRed() {
      try {
        const r = await fetch('/api/red'); const d = await r.json();
        const red = d.red;
        if (!red) { pcById('redInfo').textContent = 'No hay una red activa. Conecta la PC al punto de acceso.'; return; }
        pcById('redInfo').innerHTML = 'Adaptador: <b>' + escaparHtml(red.nombre) + '</b><br>IP actual: <b>' + red.ip + '</b> (' + (red.dhcp ? 'automática' : 'fija') + ')<br>Router: <b>' + red.puerta + '</b><br>Enlace de vendedores: <b>http://' + red.ip + ':' + d.puerto + '/vendedor</b>';
        const cfg = d.config || {};
        pcById('redOcteto').value = cfg.ultimoOcteto > 0 ? cfg.ultimoOcteto : (parseInt(red.ip.split('.')[3], 10) || '');
        pcById('redAuto').checked = !!cfg.auto;
      } catch (e) { pcById('redInfo').textContent = 'No se pudo leer la red.'; }
    }
    function redMensaje(t, ok) { const m = pcById('redMsg'); m.style.color = ok ? '#86efac' : '#fca5a5'; m.textContent = t; }
    async function fijarIpRed() {
      const octeto = parseInt(pcById('redOcteto').value, 10);
      if (!octeto) { redMensaje('Escribe el último número de la IP.', false); return; }
      redMensaje('Aplicando... (puede tardar unos segundos)', true);
      try {
        const r = await fetch('/api/red/fijar', { method: 'POST', body: JSON.stringify({ octeto, auto: pcById('redAuto').checked }) });
        const d = await r.json();
        if (!d.ok) { redMensaje(d.error || 'No se pudo fijar la IP.', false); return; }
        redMensaje((d.sinCambios ? 'La IP ya era ' : 'IP fija: ') + d.ip + '. Enlace: http://' + d.ip + ':' + location.port + '/vendedor. Si usas la bandera de Chrome en los teléfonos, agrega también ese enlace (sin /vendedor).', true);
        cargarRed();
      } catch (e) { redMensaje('Se cortó la respuesta al cambiar la IP; revisa el estado abajo.', false); setTimeout(cargarRed, 2500); }
    }
    async function ipAutomaticaRed() {
      redMensaje('Volviendo a automático...', true);
      try {
        const r = await fetch('/api/red/dhcp', { method: 'POST' }); const d = await r.json();
        redMensaje(d.ok ? 'Listo: la PC vuelve a recibir IP automática.' : (d.error || 'No se pudo.'), !!d.ok);
        setTimeout(cargarRed, 3000);
      } catch (e) { redMensaje('No se pudo contactar con el servidor.', false); }
    }

    function abrirPines() {
      document.getElementById('pinesOverlay').style.display = 'flex';
      document.getElementById('pinError').textContent = '';
      document.getElementById('pinVendedorNombre').value = '';
      document.getElementById('pinVendedorValor').value = '';
      cargarListaConPin();
      cargarEstadoClavePOS();
    }
    let claveDefinidaPOS = false;
    async function cargarEstadoClavePOS() {
      const el = document.getElementById('posClaveEstado');
      document.getElementById('posClaveAdmin').value = '';
      try {
        const r = await fetch('/api/pos/estado');
        const d = await r.json();
        claveDefinidaPOS = !!(d && d.claveDefinida);
        el.textContent = claveDefinidaPOS
          ? 'Hay una clave puesta. Escribe otra y guarda para cambiarla; guarda en blanco para quitarla.'
          : 'Todav\u00eda no hay clave: nadie puede activar el modo punto de venta.';
      } catch (e) { el.textContent = ''; }
    }
    async function guardarClavePOS() {
      const el = document.getElementById('posClaveEstado');
      const clave = document.getElementById('posClaveAdmin').value.trim();
      if (clave && (clave.length < 4 || clave.length > 20)) { el.textContent = 'La clave debe tener de 4 a 20 caracteres.'; return; }
      if (!clave && !claveDefinidaPOS) { el.textContent = 'Escribe la clave que quieres poner.'; return; }
      if ((!clave || claveDefinidaPOS) && !confirm(clave ? 'Cambiar la clave hace que todos los moviles pierdan el modo punto de venta y lo tengan que activar otra vez. Continuar?' : 'Quitar la clave hace que todos los moviles pierdan el modo punto de venta. Continuar?')) return;
      try {
        const r = await fetch('/api/pos/clave', { method: 'POST', body: JSON.stringify({ clave }) });
        const d = await r.json().catch(() => null);
        if (!r.ok || !d || !d.ok) { el.textContent = (d && d.error) || 'No se pudo guardar.'; return; }
        mostrarBanner(clave ? 'Clave de administrador guardada.' : 'Clave quitada: ya nadie tiene el modo punto de venta.');
        cargarEstadoClavePOS();
      } catch (e) { el.textContent = 'No se pudo conectar con el servidor.'; }
    }
    function cerrarPines() { document.getElementById('pinesOverlay').style.display = 'none'; }
    async function cargarListaConPin() {
      const cont = document.getElementById('listaConPin');
      try {
        const res = await fetch('/api/pines');
        const data = await res.json();
        const lista = (data && data.vendedoresConPin) || [];
        cont.innerHTML = lista.length === 0 ? '<span style="color:#64748b;">Ninguno por ahora.</span>' :
          lista.map(v => '<div style="display:flex; justify-content:space-between; padding:4px 0;"><span>' + v + '</span><button style="background:#334155; color:#fff; border:none; border-radius:6px; padding:4px 10px; font-size:12px;" onclick="quitarPin(\'' + v + '\')">Quitar</button></div>').join('');
      } catch (e) { cont.textContent = 'No se pudo cargar.'; }
    }
    async function guardarPin() {
      const vendedor = document.getElementById('pinVendedorNombre').value.trim();
      const pin = document.getElementById('pinVendedorValor').value.trim();
      const err = document.getElementById('pinError');
      err.textContent = '';
      if (!vendedor) { err.textContent = 'Escribe el nombre del vendedor.'; return; }
      if (!pin) { err.textContent = 'El PIN es obligatorio (usa el boton "Quitar" de la lista para eliminar a alguien).'; return; }
      if (!/^\d{4,6}$/.test(pin)) { err.textContent = 'El PIN debe tener de 4 a 6 numeros.'; return; }
      try {
        const res = await fetch('/api/pines', { method:'POST', body: JSON.stringify({ vendedor, pin }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) { err.textContent = (data && data.error) || 'No se pudo guardar.'; return; }
        mostrarBanner('Vendedor "' + vendedor + '" guardado con su PIN.');
        document.getElementById('pinVendedorValor').value = '';
        cargarListaConPin();
      } catch (e) { err.textContent = 'No se pudo conectar con el servidor.'; }
    }
    async function quitarPin(vendedor) {
      try {
        await fetch('/api/pines', { method:'POST', body: JSON.stringify({ vendedor, pin: '' }) });
        cargarListaConPin();
      } catch (e) {}
    }

    // ---- Nuevo pedido para vendedor (armado desde la PC) ----
    let catalogoNP = [];
    let carritoNP = [];
    let intervaloVendedoresNP = null;
    // Se recuerda que pedidos asignados ya se avisaron como "vistos" para no
    // repetir el banner cada vez que se consulta el estado.
    const asignadosVistoAvisado = {};
    // Igual, pero para el aviso de "nadie lo ha tomado todavia" (una sola vez
    // por pedido, aunque siga sin tomarse muchos minutos mas).
    const asignadosVencidoAvisado = {};

    async function revisarAsignadosVistos() {
      try {
        const res = await fetch('/api/pedidos/asignados/estado');
        const data = await res.json().catch(() => null);
        const lista = (data && data.asignados) || [];
        for (const a of lista) {
          if (a.visto && !asignadosVistoAvisado[a.id]) {
            asignadosVistoAvisado[a.id] = true;
            if (a.todos) {
              mostrarBanner(a.tomadoPor ? (a.tomadoPor + ' tomo el pedido para todos que armaste') : 'Un vendedor tomo el pedido para todos que armaste');
            } else {
              mostrarBanner((a.vendedor || 'El vendedor') + ' ya vio el pedido que le armaste');
            }
          }
          if (a.vencido && !asignadosVencidoAvisado[a.id]) {
            asignadosVencidoAvisado[a.id] = true;
            mostrarBannerDeshacer('Nadie ha tomado el pedido "para todos" #' + a.id + ' desde hace rato. ¿Lo reenvias a alguien puntual?', a.id);
          }
        }
      } catch (e) {}
    }
    setInterval(revisarAsignadosVistos, 5000);

    // Mensajes cortos sueltos que mandan los vendedores desde el movil (los que van con un pedido salen en su tarjeta).
    async function revisarMensajesPC() {
      try {
        const r = await fetch('/api/mensajes/pc');
        const d = await r.json().catch(() => null);
        const lista = (d && d.mensajes) || [];
        if (!lista.length) return;
        beep();
        mostrarBanner(lista.map(m => '\u{1F4AC} ' + m.vendedor + ': ' + m.texto).join('   |   '), 30000);
        const b = document.getElementById('banner');
        b.onclick = () => { b.style.display = 'none'; };
      } catch (e) {}
    }
    setInterval(revisarMensajesPC, 5000);
    revisarMensajesPC();

    async function cargarCatalogoNP() {
      // Se vuelve a pedir cada vez que se abre el dialogo: asi el stock que se ve (y el filtro de "sin stock") esta al dia.
      try {
        const res = await fetch('/api/catalogo');
        catalogoNP = await res.json();
      } catch (e) {}
    }

    async function cargarVendedoresNP() {
      const sel = document.getElementById('selVendedorDestino');
      const aviso = document.getElementById('npSinVendedores');
      const actual = sel.value;
      try {
        const res = await fetch('/api/vendedores');
        const data = await res.json();
        const conectados = (data && data.vendedores) || [];
        let registrados = [];
        try { const rp = await fetch('/api/pines'); registrados = ((await rp.json()) || {}).vendedoresConPin || []; } catch (e2) {}
        const todosNombres = Array.from(new Set(registrados.concat(conectados))).sort((a, b) => a.localeCompare(b));
        aviso.style.display = conectados.length === 0 ? 'block' : 'none';
        const opcionTodos = '<option value="__TODOS__">Todos los vendedores (el primero que lo tome)</option>';
        sel.innerHTML = opcionTodos + todosNombres.map(v => '<option value="' + v + '">' + v + (conectados.includes(v) ? ' \u25CF en linea' : ' (sin conexion)') + '</option>').join('');
        if (actual === '__TODOS__' || todosNombres.includes(actual)) sel.value = actual;
      } catch (e) {}
    }

    // Este mismo dialogo sirve para dos cosas: armar un pedido nuevo para un
    // vendedor (sin pedidoIdExistente), o agregar productos rapido a un
    // pedido pendiente que ya llego (pasando su id) -- ej. algo que el
    // cliente pidio despues y hay que sumarlo antes de cobrar.
    let agregarAPedidoId = null;
    let claveBaseNP = '';   // identifica esta "apertura" del dialogo: sirve para que un envio repetido no se duplique
    let enviandoNP = false;
    function hashSimpleNP(t) { let h = 5381; for (let i = 0; i < t.length; i++) h = ((h << 5) + h + t.charCodeAt(i)) | 0; return (h >>> 0).toString(36); }
    function abrirNuevoPedido(pedidoIdExistente) {
      agregarAPedidoId = pedidoIdExistente || null;
      claveBaseNP = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
      enviandoNP = false;
      document.getElementById('nuevoPedidoOverlay').style.display = 'flex';
      document.getElementById('npError').textContent = '';
      document.getElementById('npBuscador').value = '';
      document.getElementById('npResultados').innerHTML = '';
      document.getElementById('npTitulo').textContent = agregarAPedidoId ? ('Agregar al pedido #' + agregarAPedidoId) : 'Nuevo pedido para vendedor';
      document.getElementById('npBtnEnviar').textContent = agregarAPedidoId ? 'Agregar al pedido' : 'Enviar al movil del vendedor';
      document.getElementById('npDestinoWrap').style.display = agregarAPedidoId ? 'none' : 'block';
      document.getElementById('npVigenciaWrap').style.display = agregarAPedidoId ? 'none' : 'block';
      document.getElementById('npNotaWrap').style.display = agregarAPedidoId ? 'none' : 'block';
      document.getElementById('npNota').value = '';
      carritoNP = [];
      renderCarritoNP();
      cargarCatalogoNP();
      if (!agregarAPedidoId) {
        cargarVendedoresNP();
        if (!intervaloVendedoresNP) intervaloVendedoresNP = setInterval(cargarVendedoresNP, 5000);
      }
    }

    function cerrarNuevoPedido() {
      document.getElementById('nuevoPedidoOverlay').style.display = 'none';
      if (intervaloVendedoresNP) { clearInterval(intervaloVendedoresNP); intervaloVendedoresNP = null; }
      agregarAPedidoId = null;
      carritoNP = [];
      renderCarritoNP();
    }

    function buscarNP() {
      const q = normalizar(document.getElementById('npBuscador').value.trim());
      const cont = document.getElementById('npResultados');
      if (!q) { cont.innerHTML = ''; return; }
      // Si en Ajustes esta activado "Ocultar sin stock a vendedores", aqui tambien se ocultan
      // (solo cuando el Excel trae columna de cantidad, igual que en el movil del vendedor).
      const hayControlNP = catalogoNP.some(p => p.stock !== null && p.stock !== undefined);
      const baseNP = (ocultarSinStock && hayControlNP)
        ? catalogoNP.filter(p => !(p.stock === null || p.stock === undefined || p.stock <= 0))
        : catalogoNP;
      const encontrados = baseNP.filter(p =>
        normalizar(p.nombre).includes(q) || normalizar(p.sku || '').includes(q)
      ).slice(0, 25);
      cont.innerHTML = encontrados.map(p => {
        const fotoTag = tieneFoto(p.sku) ? ('<img class="np-miniatura" src="/foto/' + encodeURIComponent(p.sku) + '.jpg" onerror=\'fotoFallo(this,' + JSON.stringify(String(p.sku)) + ')\' onclick=\'verFotoProductoNP(' + JSON.stringify(p.sku) + ',' + JSON.stringify(p.nombre) + ')\'>') : '';
        return '<div class="np-resultado">' +
          fotoTag +
          '<div class="np-info"><b>' + p.nombre + '</b><span>' + (p.sku || '') + ' — $' + p.precio.toFixed(2) + '</span></div>' +
          '<button onclick=\'agregarAlCarritoNP(' + JSON.stringify(p) + ')\'>+</button>' +
        '</div>';
      }).join('');
    }

    function verFotoProductoNP(sku, nombre) {
      document.getElementById('fotoOverlayNombreNP').textContent = nombre;
      document.getElementById('fotoOverlayImgNP').src = '/foto/grande/' + encodeURIComponent(sku) + '.jpg';
      document.getElementById('fotoOverlayNP').style.display = 'flex';
    }
    function cerrarFotoProductoNP() {
      document.getElementById('fotoOverlayNP').style.display = 'none';
    }


    function agregarAlCarritoNP(p) {
      const existente = carritoNP.find(i => i.sku === p.sku);
      if (existente) existente.cantidad++;
      else carritoNP.push({ sku: p.sku, nombre: p.nombre, precio: p.precio, cantidad: 1 });
      renderCarritoNP();
      document.getElementById('npBuscador').value = '';
      document.getElementById('npResultados').innerHTML = '';
      document.getElementById('npBuscador').focus();
    }

    function cambiarCantidadNP(sku, delta) {
      const item = carritoNP.find(i => i.sku === sku);
      if (!item) return;
      item.cantidad += delta;
      if (item.cantidad <= 0) carritoNP = carritoNP.filter(i => i.sku !== sku);
      renderCarritoNP();
    }

    // Cantidad escrita a mano (ademas de los botones + y -). 0 o vacio quita el producto.
    function fijarCantidadNP(sku, valor) {
      const item = carritoNP.find(i => i.sku === sku);
      if (!item) return;
      let n = parseFloat(String(valor).replace(',', '.'));
      if (!isFinite(n) || n <= 0) {
        carritoNP = carritoNP.filter(i => i.sku !== sku);
      } else {
        n = Math.round(n * 100) / 100;
        const prod = catalogoNP.find(c => c.sku === sku);
        const errEl = document.getElementById('npError');
        if (prod && prod.stock !== null && prod.stock !== undefined && n > prod.stock) {
          n = prod.stock;
          if (errEl) errEl.textContent = 'De "' + item.nombre + '" solo hay ' + prod.stock + ' disponibles.';
        } else if (errEl) { errEl.textContent = ''; }
        if (n <= 0) carritoNP = carritoNP.filter(i => i.sku !== sku);
        else item.cantidad = n;
      }
      renderCarritoNP();
    }

    function renderCarritoNP() {
      const cont = document.getElementById('npCarrito');
      if (carritoNP.length === 0) {
        cont.innerHTML = '<div class="np-vacio">Sin productos aun.</div>';
      } else {
        cont.innerHTML = carritoNP.map(i =>
          '<div class="np-carrito-item">' +
            '<div>' + i.nombre + '</div>' +
            '<div style="display:flex; align-items:center; gap:8px;">' +
              '<button onclick="cambiarCantidadNP(\'' + i.sku + '\', -1)">-</button>' +
              '<input type="number" min="0" step="any" inputmode="decimal" value="' + i.cantidad + '" title="Escribe la cantidad" ' +
                'style="width:62px; text-align:center; padding:4px 2px; border-radius:6px; border:1px solid #475569; background:#0f172a; color:#fff; font-size:14px;" ' +
                'onfocus="this.select()" onkeydown="if(event.key===\'Enter\'){this.blur();}" onchange="fijarCantidadNP(\'' + i.sku + '\', this.value)">' +
              '<button onclick="cambiarCantidadNP(\'' + i.sku + '\', 1)">+</button>' +
            '</div>' +
          '</div>'
        ).join('');
      }
      const total = carritoNP.reduce((s, i) => s + i.precio * i.cantidad, 0);
      document.getElementById('npTotal').textContent = total.toFixed(2);
    }

    async function enviarPedidoAVendedor() {
      if (enviandoNP) return;   // doble toque: el primer envio ya esta en camino
      enviandoNP = true;
      const btnNP = document.getElementById('npBtnEnviar');
      if (btnNP) btnNP.disabled = true;
      try { await enviarPedidoAVendedorInterno(); }
      finally { enviandoNP = false; if (btnNP) btnNP.disabled = false; }
    }

    async function enviarPedidoAVendedorInterno() {
      const err = document.getElementById('npError');
      err.textContent = '';
      const notaNP = ((document.getElementById('npNota') || {}).value || '').trim();
      if (carritoNP.length === 0 && !(notaNP && !agregarAPedidoId)) { err.textContent = agregarAPedidoId ? 'Agrega al menos un producto.' : 'Agrega al menos un producto o escribe un mensaje.'; return; }
      if (agregarAPedidoId) {
        try {
          const claveAgr = claveBaseNP + '-' + hashSimpleNP(JSON.stringify(carritoNP.map(i => [i.sku, i.cantidad])));
          const res = await fetch('/api/pedidos/' + agregarAPedidoId + '/agregar', { method: 'POST', body: JSON.stringify({ items: carritoNP, claveEnvio: claveAgr }) });
          const data = await res.json().catch(() => null);
          if (!res.ok) { err.textContent = (data && data.error) || 'No se pudo agregar.'; return; }
          mostrarBanner('Se agrego al pedido #' + agregarAPedidoId + '. Se le aviso al vendedor.');
          cerrarNuevoPedido();
          cargarPedidos();
        } catch (e) { err.textContent = 'No se pudo agregar (revisa la conexion).'; }
        return;
      }
      const destino = document.getElementById('selVendedorDestino').value;
      if (!destino) { err.textContent = 'Elige a quien enviarlo.'; return; }
      const vigMin = parseInt(document.getElementById('selVigenciaNP').value, 10) || 60;
      const paraTodos = destino === '__TODOS__';
      if (carritoNP.length === 0) {
        // Sin productos: se envia solo el mensaje (le llega al movil como aviso).
        try {
          const resM = await fetch('/api/mensajes/enviar', { method: 'POST', body: JSON.stringify(paraTodos ? { todos: true, texto: notaNP } : { vendedor: destino, texto: notaNP }) });
          const dM = await resM.json().catch(() => null);
          if (!resM.ok || !dM || !dM.ok) { err.textContent = (dM && dM.error) || 'No se pudo enviar el mensaje.'; return; }
          mostrarBanner(paraTodos ? 'Mensaje enviado a todos los vendedores' : ('Mensaje enviado a ' + destino));
          cerrarNuevoPedido();
        } catch (e) { err.textContent = 'No se pudo enviar (revisa la conexion).'; }
        return;
      }
      try {
        const claveAsig = claveBaseNP + '-' + hashSimpleNP(destino + '|' + vigMin + '|' + JSON.stringify(carritoNP.map(i => [i.sku, i.cantidad])));
        const body = paraTodos ? { todos: true, items: carritoNP, minutosVigencia: vigMin, claveEnvio: claveAsig, nota: notaNP } : { vendedor: destino, items: carritoNP, minutosVigencia: vigMin, claveEnvio: claveAsig, nota: notaNP };
        const res = await fetch('/api/pedidos/asignar', { method: 'POST', body: JSON.stringify(body) });
        const data = await res.json().catch(() => null);
        if (!res.ok) { err.textContent = (data && data.error) || 'No se pudo enviar.'; return; }
        const msjEnviado = paraTodos ? ('Pedido guardado para todos los vendedores (vale ' + vigMin + ' min)') : ('Pedido guardado para ' + destino + ' (vale ' + vigMin + ' min)');
        if (data && data.id) mostrarBannerDeshacer(msjEnviado, data.id);
        else mostrarBanner(msjEnviado);
        cerrarNuevoPedido();
      } catch (e) { err.textContent = 'No se pudo enviar (revisa la conexion).'; }
    }
  </script>
</body>
</html>
'@

# ------------------------------------------------------------------
# HTML: App Vendedor (se abre en los telefonos)
# ------------------------------------------------------------------
$htmlVendedor = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no, viewport-fit=cover">
<title>Tomar Pedido - Toto Tools</title>
<link rel="manifest" href="/manifest-vendedor.json">
<link rel="icon" href="/icon-192.png">
<link rel="apple-touch-icon" href="/icon-192.png">
<meta name="theme-color" content="#2563eb">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-title" content="Pedidos Toto Tools">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<script>
  // Se aplica antes de pintar la pagina para que no haya parpadeo de claro a oscuro.
  (function() {
    var guardado = null;
    try { guardado = localStorage.getItem('temaVendedor'); } catch (e) {}
    var tema = guardado;
    if (tema !== 'oscuro' && tema !== 'claro') {
      var prefiereOscuro = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
      tema = prefiereOscuro ? 'oscuro' : 'claro';
    }
    document.documentElement.setAttribute('data-tema', tema);
  })();
</script>
<style>
  :root {
    --bg: #f1f5f9;
    --texto: #0f172a;
    --tarjeta: #fff;
    --borde: #cbd5e1;
    --borde-suave: #f1f5f9;
    --input-bg: #fff;
    --sombra: rgba(0,0,0,0.07);
    --sombra-fuerte: rgba(0,0,0,0.3);
    --secundario-texto: #475569;
    --gris-suave: #94a3b8;
    --boton-suave: #e2e8f0;
    --chip-bg: #f1f5f9;
    --mp-card-borde: #e2e8f0;
    --panel-campana-bg: #fff;
  }
  [data-tema="oscuro"] {
    --bg: #0f172a;
    --texto: #e2e8f0;
    --tarjeta: #1e293b;
    --borde: #334155;
    --borde-suave: #334155;
    --input-bg: #0f172a;
    --sombra: rgba(0,0,0,0.4);
    --sombra-fuerte: rgba(0,0,0,0.55);
    --secundario-texto: #94a3b8;
    --gris-suave: #64748b;
    --boton-suave: #334155;
    --chip-bg: #334155;
    --mp-card-borde: #334155;
    --panel-campana-bg: #1e293b;
  }
  * { box-sizing:border-box; margin:0; padding:0; font-family: 'Segoe UI', system-ui, Tahoma, Arial, sans-serif; }
  body { background:var(--bg); color:var(--texto); padding-bottom:90px; }
  header { background:#1e293b; color:#fff; padding:16px; display:flex; justify-content:space-between; align-items:center; gap:10px; }
  #headerVendedor { font-size:12px; color:#94a3b8; margin-top:3px; display:flex; align-items:center; gap:6px; }
  #btnReconectar { background:#334155; color:#fff; border:none; border-radius:12px; padding:3px 10px; font-size:11px; font-weight:700; margin-left:4px; cursor:pointer; }
  #btnReconectar:disabled { opacity:0.7; }
  body.modoCliente #btnReconectar, body.autoservicio #btnReconectar { display:none; }
  #puntoRed { width:9px; height:9px; border-radius:50%; background:#64748b; flex-shrink:0; }
  #puntoRed.ok { background:#22c55e; }
  #puntoRed.mal { background:#ef4444; }

  /* --- Diagnostico de conexion --- */
  #barrasRed { display:inline-flex; align-items:flex-end; gap:2px; height:12px; margin-left:8px; cursor:pointer; vertical-align:middle; }
  #barrasRed i { display:block; width:3px; background:#64748b; border-radius:1px; }
  #barrasRed i:nth-child(1) { height:4px; } #barrasRed i:nth-child(2) { height:6px; }
  #barrasRed i:nth-child(3) { height:9px; } #barrasRed i:nth-child(4) { height:12px; }
  #barrasRed.n1 i:nth-child(-n+1) { background:#ef4444; }
  #barrasRed.n2 i:nth-child(-n+2) { background:#f59e0b; }
  #barrasRed.n3 i:nth-child(-n+3) { background:#84cc16; }
  #barrasRed.n4 i:nth-child(-n+4) { background:#22c55e; }
  #msRed { font-size:11px; opacity:.8; margin-left:6px; cursor:pointer; }
  #bannerRed { display:none; position:sticky; top:0; z-index:44; padding:10px 14px; font-weight:700; font-size:13px; text-align:center; color:#fff; }
  #bannerRed.mal { display:block; background:#b91c1c; }
  #bannerRed.ok { display:block; background:#15803d; }
  #diagOverlay { display:none; position:fixed; inset:0; background:rgba(15,23,42,0.6); z-index:47; }
  #diagCaja { background:#f1f5f9; color:#0f172a; border-radius:0 0 16px 16px; max-height:92vh; overflow:auto; }
  .diag-fila { display:flex; justify-content:space-between; gap:10px; padding:7px 0; border-bottom:1px solid #e2e8f0; font-size:14px; }
  .diag-fila span:first-child { color:#64748b; }
  #diagSpark { display:flex; align-items:flex-end; gap:3px; height:44px; margin:10px 0; }
  #diagSpark div { flex:1; border-radius:2px; background:#22c55e; min-height:4px; }
  #diagSpark div.f { background:#ef4444; height:44px; }
  /* --- Permisos por vendedor (los pone la PC) --- */
  body.sin-ajustes #btnMenu { display:none !important; }
  body.sin-misPedidos #seccionMisPedidos { display:none !important; }
  body.sin-asignados #seccionAsignados, body.sin-asignados #bannerAsignado { display:none !important; }
  body.sin-crearPedidos #seccionPedidoActual, body.sin-crearPedidos #seccionCobro, body.sin-crearPedidos #barraTotal, body.sin-crearPedidos .resultado button { display:none !important; }
  body.sin-verStock .resultado .info > div:nth-of-type(3) { display:none !important; }
  body.sin-modoCliente #secModoCliente { display:none !important; }
  body.sin-cobrar .mp-btn-cobrar, body.sin-cancelar .mp-btn-cancelar, body.sin-editar .mp-btn-editar, body.sin-imprimir .mp-btn-imprimir { display:none !important; }
  body.sin-cobrar label:has(input[name="estadoPago"][value="cobrado"]) { display:none !important; }
  #btnMenu { position:relative; background:#334155; color:#fff; border:none; border-radius:10px; padding:10px 14px; font-size:14px; font-weight:600; cursor:pointer; white-space:nowrap; }
  #btnMenu.alerta::after { content:''; position:absolute; top:5px; right:5px; width:10px; height:10px; border-radius:50%; background:#ef4444; }
  #btnCampana { position:relative; background:#334155; color:#fff; border:none; border-radius:10px; padding:10px 12px; font-size:16px; cursor:pointer; }
  #btnTema { position:relative; background:#334155; color:#fff; border:none; border-radius:10px; padding:10px 12px; font-size:16px; cursor:pointer; }
  .badge-campana { position:absolute; top:-4px; right:-4px; background:#ef4444; color:#fff; border-radius:10px; min-width:18px; height:18px; font-size:11px; font-weight:700; display:flex; align-items:center; justify-content:center; padding:0 4px; }
  #panelCampana { position:fixed; top:66px; right:10px; left:10px; max-width:380px; margin-left:auto; background:var(--panel-campana-bg); color:var(--texto); border-radius:10px; box-shadow:0 8px 24px var(--sombra-fuerte); z-index:70; max-height:70vh; overflow:hidden; display:flex; flex-direction:column; }
  #listaCampana { overflow:auto; }
  .campana-item { padding:10px 14px; border-bottom:1px solid var(--borde-suave); font-size:13px; }
  .campana-item .campana-hora { font-size:11px; color:var(--gris-suave); margin-bottom:2px; }
  .campana-vacio { padding:16px; color:var(--gris-suave); font-size:13px; text-align:center; }
  .colapsable { width:100%; display:flex; justify-content:space-between; align-items:center; background:none; border:none; padding:0; font-size:13px; color:var(--secundario-texto); text-transform:uppercase; letter-spacing:0.6px; font-weight:700; cursor:pointer; text-align:left; }
  .colapsable .resumen { text-transform:none; letter-spacing:0; font-weight:600; font-size:12px; color:#92400e; margin-left:6px; }
  .mp-btn-items { background:none; border:none; color:#2563eb; font-size:12px; font-weight:700; padding:2px 0; margin:0 0 6px; cursor:pointer; }
  .mp-items { margin:0 0 8px; padding:8px 10px; background:var(--bg); border-radius:8px; font-size:13px; }
  .mp-items > div { display:flex; justify-content:space-between; gap:10px; padding:2px 0; }
  header h1 { font-size:18px; font-weight:600; }
  .section { background:var(--tarjeta); color:var(--texto); margin:12px; padding:16px; border-radius:12px; box-shadow:0 2px 8px var(--sombra); }
  .section h2 { font-size:13px; color:var(--secundario-texto); margin-bottom:12px; text-transform:uppercase; letter-spacing:0.6px; font-weight:700; }
  input[type=text], input[type=number], select {
    width:100%; padding:11px; border:1px solid var(--borde); border-radius:8px; font-size:15px; margin-bottom:8px; background:var(--input-bg); color:var(--texto);
  }
  .btn { background:#2563eb; color:#fff; border:none; padding:13px; border-radius:10px; font-weight:bold; font-size:15px; width:100%; cursor:pointer; }
  .btn:active { background:#1d4ed8; }
  .btn-secundario { background:#64748b; }
  #posFiltros { display:none; gap:8px; margin-top:12px; }
  #posFiltros select { flex:1; margin-bottom:0; }
  body.modoPOS #posFiltros { display:flex; }
  body.autoservicio #notaPedido, body.modoCliente #notaPedido, body.sin-mensajes #notaPedido { display:none !important; }
  body.sin-verStock #btnDescargarStock { display:none !important; }
  #estadoCatalogo { font-size:13px; color:#16a34a; margin-bottom:8px; font-weight:600; }
  #estadoCatalogo.error { color:#991b1b; }
  .resultado { display:flex; justify-content:space-between; align-items:center; padding:11px 0; border-bottom:1px solid var(--borde-suave); }
  .resultado .miniatura { width:44px; height:44px; border-radius:6px; object-fit:cover; margin-right:10px; flex-shrink:0; cursor:zoom-in; background:var(--borde-suave); }
  .resultado .info { font-size:14px; flex:1; }
  .resultado .sku { color:var(--gris-suave); font-size:12px; }
  .resultado .precio { color:var(--secundario-texto); font-size:13px; font-weight:700; }
  body.modoCliente .resultado .precio { font-size:16px; }
  .resultado button { background:var(--boton-suave); color:var(--texto); border:none; width:36px; height:36px; border-radius:8px; font-size:19px; font-weight:bold; }
  .resultado button:disabled { opacity:0.35; }
  #recientes { display:flex; flex-wrap:wrap; gap:6px; margin-top:8px; }
  #recientes .chip-reciente { background:var(--chip-bg); color:var(--texto); border:none; border-radius:16px; padding:7px 12px; font-size:12px; font-weight:600; cursor:pointer; }
  #recientes .chip-reciente:active { background:var(--boton-suave); }
  .carrito-item { display:flex; justify-content:space-between; align-items:center; padding:9px 0; border-bottom:1px solid var(--borde-suave); font-size:14px; }
  .qty-controls { display:flex; align-items:center; gap:8px; }
  .qty-controls button { width:30px; height:30px; border:none; border-radius:8px; background:var(--boton-suave); color:var(--texto); font-weight:bold; font-size:16px; }
  .qty-controls input.qty-input { width:64px; padding:6px 4px; margin:0; text-align:center; font-size:15px; font-weight:600; border:1px solid var(--borde); border-radius:8px; background:var(--input-bg); color:var(--texto); -moz-appearance:textfield; }
  .qty-controls input.qty-input::-webkit-outer-spin-button, .qty-controls input.qty-input::-webkit-inner-spin-button { -webkit-appearance:none; margin:0; }
  .total-line { font-size:18px; font-weight:bold; text-align:right; margin-top:10px; }
  .pago-opciones { display:flex; gap:14px; margin-bottom:12px; }
  .pago-opciones label { display:flex; align-items:center; gap:5px; font-size:14px; }
  #camposCobro { display:none; }
  #totalCobroTexto { font-weight:bold; margin-bottom:8px; font-size:15px; color:var(--texto); }
  #mensaje { display:none; padding:12px; border-radius:8px; margin:12px; text-align:center; font-weight:bold; }
  .ok { background:#dcfce7; color:#166534; }
  .error { background:#fee2e2; color:#991b1b; }
  #cambioTexto { font-size:13px; margin-top:4px; }
  #barraTotal { position:fixed; bottom:0; left:0; right:0; background:#1e293b; color:#fff; padding:14px 18px; display:flex; justify-content:space-between; align-items:center; font-weight:bold; font-size:16px; box-shadow:0 -3px 10px rgba(0,0,0,0.25); }
  .mp-card { border:1px solid var(--mp-card-borde); border-radius:10px; padding:10px 12px; margin-bottom:10px; }
  .mp-card .mp-top { display:flex; justify-content:space-between; font-size:13px; color:var(--secundario-texto); margin-bottom:4px; }
  .mp-card .mp-estado { font-size:11px; font-weight:bold; padding:2px 8px; border-radius:10px; }
  .mp-estado.pendiente { background:#fde047; color:#713f12; }
  .mp-estado.cobrado { background:#86efac; color:#14532d; }
  .mp-estado.cancelado { background:#e2e8f0; color:#475569; }
  .mp-card .mp-total { font-weight:bold; margin:4px 0 8px; }
  .mp-acciones { display:flex; gap:6px; flex-wrap:wrap; }
  .mp-acciones button, .mp-acciones select { flex:1; min-width:90px; padding:8px 6px; border-radius:7px; border:none; font-size:12px; font-weight:600; cursor:pointer; }
  .mp-btn-cobrar { background:#16a34a; color:#fff; }
  .mp-btn-imprimir { background:#475569; color:#fff; }
  .mp-btn-editar { background:#2563eb; color:#fff; }
  .mp-btn-cancelar { background:#b91c1c; color:#fff; }
  .mp-select { border:1px solid var(--borde); background:var(--input-bg); color:var(--texto); }
  #bannerAsignado { display:none; background:#0369a1; color:#fff; padding:12px 16px; border-radius:8px; margin:12px; font-weight:bold; text-align:center; box-shadow:0 4px 12px rgba(3,105,161,0.35); cursor:pointer; }
  #seccionAsignados { display:none; }
  .asig-card { border:1px solid #7dd3fc; background:#f0f9ff; border-radius:10px; padding:10px 12px; margin-bottom:10px; }
  .asig-card .mp-top { display:flex; justify-content:space-between; font-size:13px; color:#0369a1; margin-bottom:4px; font-weight:700; }
  .asig-card .mp-total { font-weight:bold; margin:4px 0 8px; color:#0f172a; }
  .mp-btn-visto { background:#0369a1; color:#fff; }
  .mp-btn-visto[disabled] { background:#7dd3fc; color:#0c4a6e; cursor:default; }
  .mp-btn-agregar-asig { background:#16a34a; color:#fff; }
  body.modoCliente #btnMenu, body.modoCliente #barraTotal, body.modoCliente .section:not(#seccionBuscador) { display:none !important; }
  body.modoCliente #seccionBuscador { margin-top:0; }
  body.modoCliente .resultado button, body.modoCliente .resultado .sku { display:none !important; }
  body.modoCliente #recientes { display:none !important; }
  body.modoCliente header { justify-content:center; }
  #btnSalirModoCliente { display:none; position:fixed; top:10px; right:10px; z-index:60; background:#1e293b; color:#fff; border:none; border-radius:20px; padding:10px 16px; font-size:13px; font-weight:700; }
  body.modoCliente #btnSalirModoCliente { display:block; position:static; width:calc(100% - 24px); margin:10px 12px 0; text-align:center; border-radius:12px; }
  body.autoservicio #btnMenu, body.autoservicio #btnCampana, body.autoservicio #bannerAsignado,
  body.autoservicio #seccionAsignados, body.autoservicio #seccionMisPedidos,
  body.autoservicio #accionesPostEnvio, body.autoservicio #seccionCobro .pago-opciones,
  body.autoservicio #seccionCobro h2,
  body.autoservicio #metodoPago, body.autoservicio #totalCobroTexto, body.autoservicio #camposCobro {
    display:none !important;
  }

  /* ---- Pantalla de inicio de sesion del vendedor ---- */
  #loginOverlay { --acento:#2563eb; position:fixed; inset:0; z-index:100; background:var(--bg); display:none; align-items:center; justify-content:center; padding:16px; overflow:auto; }
  [data-tema="oscuro"] #loginOverlay { --acento:#3b82f6; }
  #loginTarjeta { width:100%; max-width:400px; background:var(--tarjeta); border:1px solid var(--borde); border-radius:24px; padding:22px 20px; box-shadow:0 10px 30px var(--sombra-fuerte); margin:auto; }
  .login-marca { text-align:center; font-size:22px; font-weight:800; color:var(--acento); letter-spacing:0.3px; }
  .login-titulo { text-align:center; font-size:26px; font-weight:800; margin:2px 0 18px; color:var(--texto); }
  .login-label { display:block; font-size:15px; font-weight:700; margin:0 0 8px; color:var(--texto); }
  #loginNombreInput { width:100%; padding:15px 16px; font-size:18px; border-radius:14px; border:1px solid var(--borde); background:var(--input-bg); color:var(--texto); outline:none; }
  #loginNombreInput:focus { border-color:var(--acento); }
  .login-ayuda { font-size:14px; color:var(--secundario-texto); margin:10px 0 16px; }
  .login-usuario { display:flex; align-items:center; justify-content:space-between; gap:8px; background:var(--boton-suave); border-radius:14px; padding:10px 10px 10px 16px; margin-bottom:16px; font-size:18px; font-weight:600; color:var(--texto); }
  .login-usuario span { overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .login-usuario button { background:transparent; color:var(--texto); border:1px solid var(--borde); border-radius:10px; padding:10px 14px; font-size:14px; font-weight:600; cursor:pointer; flex-shrink:0; }
  #loginPuntos { display:flex; justify-content:center; gap:16px; margin:6px 0 16px; }
  #loginPuntos span { width:16px; height:16px; border-radius:50%; border:2px solid var(--gris-suave); transition:background .12s, border-color .12s; }
  #loginPuntos span.lleno { background:var(--acento); border-color:var(--acento); }
  #loginTeclado { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:16px; }
  #loginTeclado button { height:clamp(48px,8vh,60px); font-size:26px; font-weight:600; border-radius:16px; border:1px solid transparent; background:var(--boton-suave); color:var(--texto); cursor:pointer; -webkit-tap-highlight-color:transparent; }
  #loginTeclado button:active { background:var(--acento); color:#fff; }
  #loginTeclado button.aux { background:transparent; border-color:var(--borde); color:var(--secundario-texto); }
  #loginOverlay .btn { background:var(--acento); padding:15px; font-size:17px; margin-top:0; }
  #btnLoginEntrar:disabled { opacity:0.4; cursor:default; }
  .login-olvido-link { display:block; text-align:center; margin-top:14px; font-size:14px; color:var(--secundario-texto); text-decoration:underline; }
  #loginOlvido { display:none; margin-top:10px; font-size:13px; color:var(--secundario-texto); text-align:center; background:var(--boton-suave); border-radius:10px; padding:10px; }
  #loginError { min-height:18px; text-align:center; color:#ef4444; font-size:13px; font-weight:600; margin-top:10px; }
  @keyframes loginSacudir { 0%,100%{transform:translateX(0)} 20%{transform:translateX(-9px)} 40%{transform:translateX(9px)} 60%{transform:translateX(-6px)} 80%{transform:translateX(6px)} }
  #loginTarjeta.sacudir { animation:loginSacudir .35s; }
  /* --- Sin zoom y ajustado a la pantalla del movil --- */
  html { -webkit-text-size-adjust:100%; text-size-adjust:100%; touch-action:manipulation; overflow-x:hidden; }
  body { max-width:100vw; overflow-x:hidden; touch-action:pan-x pan-y; -webkit-tap-highlight-color:transparent; }
  img, table, textarea, select, input { max-width:100%; }
  input, select, textarea { font-size:16px !important; }
  #menuOverlay > div, #diagCaja { max-width:100vw; }
</style>
<script>
document.addEventListener('gesturestart', function (e) { e.preventDefault(); });
document.addEventListener('gesturechange', function (e) { e.preventDefault(); });
document.addEventListener('touchmove', function (e) { if (e.touches && e.touches.length > 1) e.preventDefault(); }, { passive: false });
</script>
</head>
<body>
  <header>
    <div><h1>Toto Tools - Tomar Pedido</h1><div id="headerVendedor"><span id="puntoRed"></span><span id="textoRed">Conectando...</span><button id="btnReconectar" type="button" onclick="reconectarAhora()">&#8635; Reconectar</button></div></div>
    <div style="display:flex; align-items:center; gap:8px;">
      <button id="btnTema" onclick="cambiarTema()" title="Cambiar tema">&#127769;</button>
      <button id="btnCampana" onclick="toggleCampana()">&#128276;<span id="badgeCampana" class="badge-campana" style="display:none;">0</span></button>
      <button id="btnMenu" onclick="abrirMenu()">&#9881; Ajustes</button>
    </div>
  </header>

  <div id="bannerRed" onclick="abrirDiagnostico()"></div>

  <div id="panelCampana" style="display:none;">
    <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 14px; background:#1e293b; color:#fff; border-radius:10px 10px 0 0;">
      <strong style="font-size:14px;">Notificaciones</strong>
      <button onclick="toggleCampana()" style="background:#334155; color:#fff; border:none; border-radius:7px; padding:5px 10px; font-size:12px;">Cerrar</button>
    </div>
    <div id="listaCampana"></div>
  </div>

  <div id="colaOfflineAviso" style="display:none; margin:12px; padding:10px 14px; border-radius:8px; background:#fef3c7; color:#92400e; font-size:13px; font-weight:600;"></div>

  <div id="bannerAsignado" onclick="irAAsignados()">La caja te armo un pedido nuevo — toca para verlo</div>

  <button id="btnSalirModoCliente" onclick="salirModoCliente()">&#10005; Salir del modo cliente</button>

  <div class="section" id="seccionAsignados">
    <h2>Pedidos que te armo la caja</h2>
    <div id="asignadosLista"></div>
  </div>

  <div class="section" id="seccionBuscador">
    <h2>Buscar producto</h2>
    <input type="text" id="buscador" placeholder="Nombre o SKU" oninput="buscar()">
    <div id="resultados"></div>
    <div id="recientes"></div>
  </div>

  <div class="section" id="seccionPedidoActual">
    <h2>Pedido actual</h2>
    <div id="carrito"></div>
    <div class="total-line">Total: $<span id="totalPedido">0.00</span><span id="notaTransferenciaTotal" style="display:none; font-size:12px; font-weight:normal; color:#0369a1;"> (transferencia x2)</span></div>
  </div>

  <div class="section" id="seccionCobro">
    <h2>Cobro</h2>
    <div class="pago-opciones">
      <label><input type="radio" name="estadoPago" value="pendiente" checked onchange="toggleCobro()"> Pendiente de pago</label>
      <label><input type="radio" name="estadoPago" value="cobrado" onchange="toggleCobro()"> Cobrado</label>
    </div>
    <select id="metodoPago" onchange="guardarMetodoPagoPreferido(); renderCarrito();">
      <option value="Efectivo">Efectivo</option>
      <option value="Transferencia">Transferencia (precio x2)</option>
      <option value="Combinado">Efectivo + Transferencia</option>
      <option value="Otro">Otro</option>
    </select>
    <div id="totalCobroTexto"></div>
    <div id="camposCobro">
      <input type="number" id="montoRecibido" placeholder="Monto recibido" oninput="calcularCambio()">
      <div id="camposCobroCombinado" style="display:none;">
        <input type="number" id="montoEfectivoCombo" placeholder="Parte en efectivo (del pedido)" oninput="calcularCambio()">
        <input type="number" id="montoTransferCombo" placeholder="Parte en transferencia (del pedido, se cobra x2)" oninput="calcularCambio()">
      </div>
      <div id="cambioTexto"></div>
    </div>
    <input type="text" id="notaPedido" maxlength="200" placeholder="Nota para la caja (opcional)" oninput="actualizarBotonEnviar()">
    <button class="btn" id="btnEnviarPedido" onclick="enviarPedido()">Enviar pedido</button>
  </div>

  <div class="section" id="seccionMisPedidos">
    <button class="colapsable" onclick="toggleMisPedidos()">
      <span><span id="tituloMisPedidos">Mis pedidos de hoy</span><span class="resumen" id="resumenMisPedidos"></span></span>
      <span id="flechaMisPedidos">&#9662;</span>
    </button>
    <div id="posFiltros">
      <select id="posRango" onchange="posUltimaFirma = null; cargarVentasPOS()"><option value="1">Hoy</option><option value="7">7 d&iacute;as</option><option value="30">30 d&iacute;as</option></select>
      <select id="posQuien" onchange="posUltimaFirma = null; cargarVentasPOS()"><option value="todos">Todos</option><option value="yo">Solo yo</option></select>
    </div>
    <div id="misPedidos" style="display:none; margin-top:12px;"><div style="color:#94a3b8; font-size:13px;">Sin pedidos todavia hoy.</div></div>
  </div>

  <div id="mensaje"></div>
  <div id="accionesPostEnvio" style="display:none; margin:0 12px 12px;">
    <button class="btn btn-secundario" onclick="imprimirUltimoPedido()">Imprimir recibo</button>
  </div>
  <div id="barraTotal">
    <span>Total del pedido</span>
    <div style="text-align:right;">
      <span id="totalPedidoSticky">$0.00</span>
      <div id="totalUSDSticky" style="display:none; font-size:11px; font-weight:normal; color:#94a3b8; line-height:1.3;"></div>
    </div>
  </div>

  <div id="loginOverlay">
    <div id="loginTarjeta">
      <div class="login-marca">Toto Tools</div>
      <div class="login-titulo">Iniciar sesi&oacute;n</div>

      <div id="loginPasoNombre">
        <label class="login-label" for="loginNombreInput">Usuario</label>
        <input type="text" id="loginNombreInput" list="dlUsuariosLogin" placeholder="Escribe o elige tu nombre" autocomplete="off" autocapitalize="words" onkeydown="if (event.key === 'Enter') loginContinuar()">
        <datalist id="dlUsuariosLogin"></datalist>
        <p class="login-ayuda">Elige tu nombre de la lista (o escribelo) y entra con tu PIN</p>
        <button type="button" class="btn" onclick="loginContinuar()">Continuar</button>
      </div>

      <div id="loginPasoPin" style="display:none;">
        <label class="login-label">Usuario</label>
        <div class="login-usuario"><span id="loginNombreTxt"></span><button type="button" onclick="mostrarPasoNombre()">Cambiar</button></div>
        <label class="login-label">PIN</label>
        <div id="loginPuntos"><span></span><span></span><span></span><span></span><span></span><span></span></div>
        <div id="loginTeclado">
          <button type="button" onclick="loginTecla('1')">1</button><button type="button" onclick="loginTecla('2')">2</button><button type="button" onclick="loginTecla('3')">3</button><button type="button" onclick="loginTecla('4')">4</button><button type="button" onclick="loginTecla('5')">5</button><button type="button" onclick="loginTecla('6')">6</button><button type="button" onclick="loginTecla('7')">7</button><button type="button" onclick="loginTecla('8')">8</button><button type="button" onclick="loginTecla('9')">9</button>
          <button type="button" class="aux" onclick="loginTecla('C')">C</button>
          <button type="button" onclick="loginTecla('0')">0</button>
          <button type="button" class="aux" onclick="loginTecla('B')">&#9003;</button>
        </div>
        <button type="button" class="btn" id="btnLoginEntrar" onclick="loginEntrar()" disabled>Entrar</button>
        <a href="#" class="login-olvido-link" onclick="document.getElementById('loginOlvido').style.display='block'; return false;">&iquest;Olvidaste tu PIN?</a>
        <div id="loginOlvido">P&iacute;dele al encargado que te quite o cambie el PIN desde la PC (men&uacute;, &quot;PIN de vendedores&quot;).</div>
      </div>

      <div id="loginError"></div>
    </div>
  </div>

  <div id="menuOverlay" style="display:none; position:fixed; inset:0; background:rgba(15,23,42,0.55); z-index:45;" onclick="if (event.target === this) cerrarMenu()">
    <div style="background:#f1f5f9; border-radius:0 0 16px 16px; max-height:92vh; overflow:auto; padding-bottom:6px;">
      <div style="display:flex; justify-content:space-between; align-items:center; padding:14px 16px; background:#1e293b; color:#fff;">
        <strong style="font-size:16px;">Ajustes</strong>
        <button onclick="cerrarMenu()" style="background:#334155; color:#fff; border:none; border-radius:8px; padding:9px 16px; font-weight:600; font-size:14px;">Cerrar</button>
      </div>

      <div class="section">
        <h2>Vendedor</h2>
        <input type="text" id="nombreVendedor" placeholder="Tu nombre">
        <label style="font-size:12px; color:#64748b; display:block; margin:2px 0 4px;">Tu PIN (obligatorio, lo tuvo que crear el encargado en la PC; para cambiarlo escribe primero el actual):</label>
        <input type="text" id="pinVendedor" inputmode="numeric" maxlength="6" placeholder="Tu PIN (4 a 6 numeros)">
        <div id="estadoPinVendedor" style="font-size:12px; color:#64748b; margin-top:4px;"></div>
        <button class="btn btn-secundario" style="margin-top:10px;" onclick="cerrarSesion()">Cerrar sesion</button>
      </div>

      <div class="section">
        <h2>Catalogo de productos</h2>
        <div id="estadoCatalogo">Cargando catalogo...</div>
        <button class="btn btn-secundario" onclick="cargarCatalogo()">Actualizar catalogo</button>
        <button class="btn btn-secundario" style="margin-top:8px;" onclick="document.getElementById('archivoExcelMovil').click()">Subir Excel desde el movil</button>
        <input type="file" id="archivoExcelMovil" accept=".xlsx,.xls" style="display:none">
        <div id="estadoSubidaExcel" style="font-size:12px; color:#64748b; margin-top:6px;"></div>
        <button class="btn btn-secundario" style="margin-top:8px;" onclick="descargarExcelCatalogo()">Descargar el Excel de la PC</button>
        <button class="btn btn-secundario" id="btnDescargarStock" style="margin-top:8px;" onclick="descargarStockActual()">Descargar lo que va quedando (stock actual)</button>
        <div id="estadoDescargaExcel" style="font-size:12px; color:#64748b; margin-top:6px;"></div>
      </div>

      <div class="section">
        <h2>Fotos del catalogo</h2>
        <p style="font-size:12px; color:#64748b; margin-bottom:8px;">En la app del catalogo, toca "Copia de seguridad" (baja un .zip a Descargas) y luego elige ese archivo aqui. Las fotos quedan guardadas en este mismo telefono (no se mandan a la PC ni se gasta WiFi); se emparejan con los productos por SKU.</p>
        <button class="btn btn-secundario" onclick="document.getElementById('zipFotosMovil').click()">Cargar fotos en este telefono</button>
        <input type="file" id="zipFotosMovil" accept=".zip" style="display:none">
        <div id="estadoSubidaFotos" style="font-size:12px; color:#64748b; margin-top:6px;"></div>
      </div>

      <div class="section">
        <h2>Avisos en este teléfono</h2>
        <p style="font-size:12px; color:#64748b; margin-bottom:8px;">La caja te avisa si cambia un precio, se agota un producto, te anula un pedido o cambia tus permisos. Con la app abierta suena y vibra siempre.</p>
        <button class="btn btn-secundario" onclick="activarAvisosTelefono()">Activar avisos del teléfono</button>
        <div id="estadoAvisosTel" style="font-size:12px; color:#64748b; margin-top:6px;"></div>
      </div>

      <div class="section" id="secRedSenal">
        <h2>Conexi&oacute;n con la PC</h2>
        <div style="display:flex; align-items:center; gap:14px; margin:6px 0 4px;">
          <span id="barrasRed" class="n0" onclick="abrirDiagnostico()" style="transform:scale(1.7); transform-origin:left bottom; margin:0 22px 4px 4px;"><i></i><i></i><i></i><i></i></span>
          <span id="msRed" onclick="abrirDiagnostico()" style="font-size:14px; color:#64748b;"></span>
        </div>
        <button class="btn btn-secundario" onclick="abrirDiagnostico()">Ver diagn&oacute;stico de la se&ntilde;al</button>
      </div>

      <div class="section" id="secPOS">
        <h2>Modo punto de venta</h2>
        <div id="posInactivoBox">
          <p style="font-size:12px; color:#64748b; margin-bottom:8px;">Lo activa el administrador con su clave: este tel&eacute;fono cobra al enviar, permite descuentos y muestra el historial de ventas, las devoluciones y el reporte de efectivo y transferencia. Sigue sincronizado con la PC.</p>
          <input type="password" id="posClaveInput" placeholder="Clave de administrador" autocomplete="off" onkeydown="if (event.key === 'Enter') activarModoPOS()">
          <button class="btn btn-secundario" onclick="activarModoPOS()">Activar modo punto de venta</button>
        </div>
        <div id="posActivoBox" style="display:none;">
          <p style="font-size:13px; color:#166534; font-weight:600; margin-bottom:8px;">Modo punto de venta ACTIVO en este tel&eacute;fono. Mira "Ventas (caja)" en la pantalla principal.</p>
          <button class="btn btn-secundario" onclick="desactivarModoPOS()">Desactivar en este tel&eacute;fono</button>
        </div>
        <div id="posMensaje" style="font-size:12px; color:#991b1b; margin-top:6px;"></div>
      </div>

      <div class="section" id="secModoCliente">
        <h2>Catalogo para mostrarle al cliente</h2>
        <p style="font-size:12px; color:#64748b; margin-bottom:8px;">Oculta el stock, los botones de cobro y "Mis pedidos": el cliente solo ve nombres y precios en el buscador. Toca la X arriba a la derecha para salir.</p>
        <button class="btn btn-secundario" onclick="activarModoCliente()">Mostrarle el catalogo al cliente</button>
      </div>
    </div>
  </div>

  <div id="configOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.6); align-items:center; justify-content:center; z-index:50;">
    <div style="background:#fff; padding:20px; border-radius:12px; max-width:420px; width:92%; max-height:85vh; overflow:auto;">
      <h2 style="font-size:16px; margin-bottom:6px;">Mapear columnas del Excel</h2>
      <p style="font-size:12px; color:#64748b; margin-bottom:10px;">Elige que columna corresponde a cada dato (SKU y Cantidad/Stock son opcionales):</p>
      <div id="mapeoCampos"></div>
      <button class="btn" onclick="guardarMapeoYCerrar()" style="margin-top:10px;">Guardar y actualizar catalogo</button>
      <button class="btn btn-secundario" style="margin-top:6px;" onclick="cerrarConfig()">Cancelar</button>
      <div id="configPaso2Error" style="color:#991b1b; font-size:12px; margin-top:8px;"></div>
    </div>
  </div>

  <div id="diagOverlay" onclick="if (event.target === this) cerrarDiagnostico()">
    <div id="diagCaja">
      <div style="display:flex; justify-content:space-between; align-items:center; padding:14px 16px; background:#1e293b; color:#fff;">
        <strong style="font-size:16px;">Diagnóstico de conexión</strong>
        <button onclick="cerrarDiagnostico()" style="background:#334155; color:#fff; border:none; border-radius:8px; padding:9px 16px; font-weight:600; font-size:14px;">Cerrar</button>
      </div>
      <div class="section">
        <div id="diagResumen" style="font-size:20px; font-weight:700; margin-bottom:4px;">Comprobando...</div>
        <div id="diagSpark"></div>
        <div id="diagDetalle"></div>
        <div id="diagConsejo" style="font-size:13px; color:#334155; background:#e2e8f0; border-radius:8px; padding:10px 12px; margin-top:10px;"></div>
        <button class="btn" id="diagBtnProbar" style="margin-top:12px;" onclick="probarAhora()">Probar ahora</button>
        <p style="font-size:11px; color:#64748b; margin-top:8px;">La calidad se mide con la rapidez y las fallas al hablar con la PC del negocio (no con internet).</p>
      </div>
    </div>
  </div>

  <div id="fotoOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.85); align-items:center; justify-content:center; z-index:60;" onclick="cerrarFotoProducto()">
    <div style="max-width:94vw; max-height:90vh; text-align:center;">
      <img id="fotoOverlayImg" src="" style="max-width:94vw; max-height:80vh; border-radius:10px; background:#fff;">
      <div id="fotoOverlayNombre" style="color:#fff; margin-top:10px; font-size:15px;"></div>
    </div>
  </div>

  <script src="/xlsx.js"></script>
  <script>
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('/sw.js').catch(() => {});
    }

    // ---- Cada peticion lleva el nombre del vendedor (la PC lo usa para aplicar sus permisos) ----
    (function () {
      const fetchOriginal = window.fetch.bind(window);
      window.fetch = function (input, init) {
        try {
          const campo = document.getElementById('nombreVendedor');
          const n = campo ? (campo.value || '').trim() : '';
          if (n && new URLSearchParams(location.search).get('cliente') !== '1' && typeof input === 'string' && input.charAt(0) === '/') {
            init = Object.assign({}, init);
            const h = new Headers(init.headers || {});
            if (!h.has('X-Vendedor')) h.set('X-Vendedor', encodeURIComponent(n));
            const tokPos = localStorage.getItem('posToken');
            if (tokPos && !h.has('X-Pos-Token')) h.set('X-Pos-Token', tokPos);
            init.headers = h;
          }
        } catch (e) {}
        return fetchOriginal(input, init);
      };
    })();

    // ---- Tema oscuro/claro (automatico segun el telefono, o manual) ----
    function aplicarTema(tema) {
      document.documentElement.setAttribute('data-tema', tema);
      const btn = document.getElementById('btnTema');
      if (btn) btn.innerHTML = tema === 'oscuro' ? '&#9728;&#65039;' : '&#127769;';
    }
    function cambiarTema() {
      const actual = document.documentElement.getAttribute('data-tema') === 'oscuro' ? 'oscuro' : 'claro';
      const nuevo = actual === 'oscuro' ? 'claro' : 'oscuro';
      aplicarTema(nuevo);
      try { localStorage.setItem('temaVendedor', nuevo); } catch (e) {}
    }
    // El tema inicial ya lo puso el script del <head> (para evitar parpadeo);
    // aqui solo se ajusta el icono del boton a lo que quedo aplicado.
    aplicarTema(document.documentElement.getAttribute('data-tema') === 'oscuro' ? 'oscuro' : 'claro');
    // Si el vendedor nunca eligio manualmente, seguir el tema del telefono en vivo
    // (por ejemplo si el telefono pasa a modo oscuro solo de noche).
    if (window.matchMedia) {
      window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
        let guardado = null;
        try { guardado = localStorage.getItem('temaVendedor'); } catch (err) {}
        if (guardado !== 'oscuro' && guardado !== 'claro') aplicarTema(e.matches ? 'oscuro' : 'claro');
      });
    }

    let catalogo = [];
    let carrito = [];
    // Cuando un pedido "para todos" que tomaste se agrega al carrito, se anota
    // aqui quien lo tomo y cuando, para mandarlo junto con el pedido y que
    // quede visible en la PC cuando se cobre (historial de quien atendio).
    let origenesAsignadosCarrito = [];
    let ocultarSinStock = false;   // lo decide la PC (/api/config)
    let tasaDolarActual = 0;       // idem, tasa USD configurada en la PC (0 = no mostrar)
    let permitirDescuentosActual = false; // idem, lo activa la PC en Ajustes (por defecto no se pueden editar precios)
    let umbralStockBajoActual = 0; // idem, marca en el buscador cuando queda poco (0 = desactivado)
    let autoservicioDestino = 'pc';   // idem, decide adonde caen los pedidos de autoservicio (/api/config)

    // Modo autoservicio: se activa entrando con ?cliente=1 en el enlace (el
    // vendedor o la caja comparte ese enlace con los clientes). El cliente ve
    // el catalogo completo y arma su propio pedido, sin acceso a Ajustes ni a
    // las secciones de vendedor.
    const autoservicioActivo = new URLSearchParams(location.search).get('cliente') === '1';
    if (autoservicioActivo) document.body.classList.add('autoservicio');

    // Si el Excel no tiene columna de cantidad (ningun producto trae stock),
    // no se puede saber que esta agotado: en ese caso no se oculta nada.
    function hayControlStock() {
      return catalogo.some(p => p.stock !== null && p.stock !== undefined);
    }
    function productoSinStock(p) {
      return p.stock === null || p.stock === undefined || p.stock <= 0;
    }
    function ocultandoSinStock() {
      return ocultarSinStock && hayControlStock();
    }

    // ---- Vendedor ----
    const nombreInput = document.getElementById('nombreVendedor');
    // En autoservicio (enlace/QR) el nombre del cliente se guarda aparte, para
    // no heredar el nombre de un vendedor que use ese mismo telefono.
    const claveNombreLocal = autoservicioActivo ? 'clienteNombre' : 'vendedorNombre';
    nombreInput.value = localStorage.getItem(claveNombreLocal) || '';
    nombreInput.addEventListener('input', () => { localStorage.setItem(claveNombreLocal, nombreInput.value); pintarVendedorHeader(); });
    const pinInput = document.getElementById('pinVendedor');
    pinInput.value = localStorage.getItem('vendedorPin') || '';
    let pinVendedorAlEnfocar = pinInput.value;
    pinInput.addEventListener('focus', () => { pinVendedorAlEnfocar = pinInput.value; });
    pinInput.addEventListener('input', () => localStorage.setItem('vendedorPin', pinInput.value.trim()));
    pinInput.addEventListener('blur', sincronizarPinVendedor);
    function miPin() { return (localStorage.getItem('vendedorPin') || '').trim(); }

    // El vendedor pone o cambia su propio PIN desde aqui mismo (sin pasar
    // por la PC). Si ya tenia uno puesto, hace falta el PIN actual correcto
    // para cambiarlo (se manda el que habia al entrar al campo).
    async function sincronizarPinVendedor() {
      const vendedor = (nombreInput.value || '').trim();
      const pinNuevo = pinInput.value.trim();
      const estado = document.getElementById('estadoPinVendedor');
      if (!vendedor || pinNuevo === pinVendedorAlEnfocar) { if (estado) estado.textContent = ''; return; }
      try {
        const res = await fetch('/api/pines', { method: 'POST', body: JSON.stringify({ vendedor, pin: pinNuevo, pinActual: pinVendedorAlEnfocar }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          if (estado) estado.textContent = (data && data.error) || 'No se pudo guardar el PIN en la PC.';
          pinInput.value = pinVendedorAlEnfocar;
          localStorage.setItem('vendedorPin', pinVendedorAlEnfocar);
        } else if (estado) {
          estado.textContent = pinNuevo ? 'PIN guardado.' : 'PIN quitado.';
          pinVendedorAlEnfocar = pinNuevo;
        }
      } catch (e) {
        if (estado) estado.textContent = 'Sin conexion con la PC: se reintenta la proxima vez que toques este campo.';
      }
    }

    // ---- Menu de ajustes (nombre, catalogo, Excel): fuera de la pantalla del pedido ----
    function pintarVendedorHeader() {
      const t = (nombreInput.value || '').trim();
      document.getElementById('textoRed').textContent = autoservicioActivo
        ? (t ? ('Cliente: ' + t) : 'Cliente')
        : (t ? ('Vendedor: ' + t) : 'Sin sesion: inicia sesion');
    }

    // ---- Diagnostico de conexion con la PC (calidad de la senal + aviso si el servidor no responde) ----
    const diagRed = { muestras: [], fallos: 0, ultimoOk: 0, caido: false, hora: '', probando: false };

    function medirPing() {
      const intento = new Promise(resolve => {
        const ini = performance.now();
        fetch('/api/ping?t=' + Date.now(), { cache: 'no-store' })
          .then(r => { if (!r.ok) throw new Error('http ' + r.status); return r.json(); })
          .then(d => resolve({ ok: true, ms: Math.round(performance.now() - ini), hora: (d && d.hora) || '' }))
          .catch(() => resolve({ ok: false, ms: 0 }));
      });
      const tope = new Promise(resolve => setTimeout(() => resolve({ ok: false, ms: 0 }), 4500));
      return Promise.race([intento, tope]);
    }

    function calidadRed() {
      const m = diagRed.muestras;
      if (m.length === 0) return { barras: 0, texto: 'Comprobando...', media: 0, max: 0, perdida: 0 };
      const ok = m.filter(x => x.ok);
      const perdida = Math.round(100 * (m.length - ok.length) / m.length);
      const media = ok.length ? Math.round(ok.reduce((s, x) => s + x.ms, 0) / ok.length) : 0;
      const max = ok.length ? Math.max.apply(null, ok.map(x => x.ms)) : 0;
      let barras, texto;
      if (ok.length === 0 || diagRed.fallos >= 2) { barras = 0; texto = 'Sin conexión con la PC'; }
      else if (perdida >= 30 || media > 900) { barras = 1; texto = 'Señal mala'; }
      else if (perdida > 0 || media > 350) { barras = 2; texto = 'Señal regular'; }
      else if (media > 120) { barras = 3; texto = 'Señal buena'; }
      else { barras = 4; texto = 'Señal excelente'; }
      return { barras, texto, media, max, perdida };
    }

    function mostrarBannerRed(recuperada) {
      const b = document.getElementById('bannerRed');
      if (!b) return;
      if (recuperada) {
        b.className = 'ok';
        b.textContent = '✔ Conexión con la PC restablecida';
        setTimeout(() => { if (!diagRed.caido) b.className = ''; }, 4000);
      } else {
        b.className = 'mal';
        b.textContent = navigator.onLine === false
          ? '⚠ El teléfono no está conectado al Wi‑Fi. Tus pedidos se guardan y se envían al volver.'
          : '⚠ La PC no responde. Revisa el Wi‑Fi o que el servidor siga abierto. Tus pedidos se guardan y se envían al volver. (toca para ver el diagnóstico)';
      }
    }

    function registrarMuestraRed(m) {
      diagRed.muestras.push(m);
      if (diagRed.muestras.length > 12) diagRed.muestras.shift();
      if (m.ok) { diagRed.fallos = 0; diagRed.ultimoOk = Date.now(); if (m.hora) diagRed.hora = m.hora; }
      else diagRed.fallos++;
      if (!m.ok && diagRed.fallos >= 2 && !diagRed.caido) {
        diagRed.caido = true;
        mostrarBannerRed(false);
        try { beepAsignado(); } catch (e) {}
        try { if (navigator.vibrate) navigator.vibrate([300, 150, 300]); } catch (e) {}
        try { mostrarNotificacionSistema('Sin conexión con la PC', 'El servidor de pedidos no responde.'); } catch (e) {}
      } else if (m.ok && diagRed.caido) {
        diagRed.caido = false;
        mostrarBannerRed(true);
        try { intentarEnviarCola(); } catch (e) {}
        try { cargarCatalogo(); } catch (e) {}
      }
      pintarRed();
    }

    function pintarRed() {
      const q = calidadRed();
      const punto = document.getElementById('puntoRed');
      const barras = document.getElementById('barrasRed');
      const ms = document.getElementById('msRed');
      if (punto) {
        punto.className = q.barras === 0 && diagRed.muestras.length ? 'mal' : (q.barras >= 3 ? 'ok' : '');
        punto.title = q.texto + (q.media ? (' (' + q.media + ' ms)') : '');
      }
      if (barras) barras.className = 'n' + q.barras;
      if (ms) ms.textContent = q.barras === 0 ? (diagRed.muestras.length ? 'sin señal' : '') : (q.media + ' ms');
      const ov = document.getElementById('diagOverlay');
      if (ov && ov.style.display === 'block') pintarDiagnostico();
    }

    async function revisarConexion() {
      registrarMuestraRed(await medirPing());
    }

    function abrirDiagnostico() {
      document.getElementById('diagOverlay').style.display = 'block';
      pintarDiagnostico();
      probarAhora();
    }
    function cerrarDiagnostico() { document.getElementById('diagOverlay').style.display = 'none'; }

    async function probarAhora() {
      if (diagRed.probando) return;
      diagRed.probando = true;
      const btn = document.getElementById('diagBtnProbar');
      if (btn) { btn.disabled = true; btn.textContent = 'Probando...'; }
      for (let i = 0; i < 5; i++) {
        registrarMuestraRed(await medirPing());
        await new Promise(r => setTimeout(r, 250));
      }
      diagRed.probando = false;
      if (btn) { btn.disabled = false; btn.textContent = 'Probar ahora'; }
    }

    function pintarDiagnostico() {
      const q = calidadRed();
      const colores = ['#b91c1c', '#dc2626', '#d97706', '#65a30d', '#16a34a'];
      const r = document.getElementById('diagResumen');
      r.textContent = q.texto;
      r.style.color = colores[q.barras];
      document.getElementById('diagSpark').innerHTML = diagRed.muestras.map(m =>
        m.ok ? ('<div style="height:' + Math.min(44, 6 + m.ms / 8) + 'px" title="' + m.ms + ' ms"></div>') : '<div class="f" title="sin respuesta"></div>'
      ).join('');
      const hace = diagRed.ultimoOk ? Math.round((Date.now() - diagRed.ultimoOk) / 1000) : null;
      const c = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
      const filas = [
        ['Wi‑Fi / red del teléfono', navigator.onLine === false ? 'Desconectado' : 'Conectado'],
        ['Respuesta media de la PC', q.media ? (q.media + ' ms') : '—'],
        ['Respuesta más lenta', q.max ? (q.max + ' ms') : '—'],
        ['Pruebas sin respuesta', q.perdida + ' %'],
        ['Última respuesta de la PC', hace === null ? 'ninguna' : ('hace ' + hace + ' s')],
        ['Hora de la PC', diagRed.hora || '—']
      ];
      if (c && (c.effectiveType || c.rtt || c.downlink)) {
        filas.push(['Datos del teléfono (orientativo)', [c.effectiveType, c.downlink ? (c.downlink + ' Mbps') : '', c.rtt ? (c.rtt + ' ms') : ''].filter(Boolean).join(' · ')]);
      }
      document.getElementById('diagDetalle').innerHTML = filas.map(f => '<div class="diag-fila"><span>' + f[0] + '</span><span><b>' + f[1] + '</b></span></div>').join('');
      let consejo;
      if (navigator.onLine === false) consejo = 'El teléfono no tiene Wi‑Fi. Actívalo y conéctate a la red del negocio.';
      else if (q.barras === 0 && diagRed.muestras.length) consejo = 'El teléfono tiene red pero la PC no contesta. Revisa que la PC esté encendida, que la ventana del servidor siga abierta y que estés en la misma red Wi‑Fi. Si nada cambia, el router puede estar aislando los dispositivos.';
      else if (q.barras <= 2) consejo = 'La señal con la PC es floja. Acércate al router o a la PC y evita paredes gruesas o metal entre ambos.';
      else consejo = 'Todo bien: la conexión con la PC es estable.';
      document.getElementById('diagConsejo').textContent = consejo;
    }

    // Revision de la conexion: cada 6 s normalmente, pero cada 2 s mientras
    // esta caida (asi al volver al alcance del WiFi se recupera enseguida), y
    // al instante cuando el telefono recupera red o se vuelve a abrir la app.
    async function revisarConexionYProgramar() {
      try { await revisarConexion(); } catch (e) {}
      setTimeout(revisarConexionYProgramar, diagRed.caido ? 2000 : 6000);
    }
    revisarConexionYProgramar();
    window.addEventListener('online', revisarConexion);
    window.addEventListener('offline', revisarConexion);
    document.addEventListener('visibilitychange', () => { if (!document.hidden) revisarConexion(); });
    window.addEventListener('focus', revisarConexion);

    // Boton "Reconectar": prueba la conexion con la PC ya mismo (hasta 4
    // intentos) y, si responde, manda lo pendiente y refresca todo sin esperar
    // a los avisos automaticos.
    let reconectando = false;
    async function reconectarAhora() {
      if (reconectando) return;
      reconectando = true;
      const b = document.getElementById('btnReconectar');
      const poner = (txt, dis) => { if (b) { b.textContent = txt; b.disabled = !!dis; } };
      poner('Conectando...', true);
      let ok = false;
      for (let i = 0; i < 4 && !ok; i++) {
        const m = await medirPing();
        registrarMuestraRed(m);
        ok = m.ok;
        if (!ok) await new Promise(r => setTimeout(r, 1000));
      }
      if (ok) {
        const tareas = [intentarEnviarCola, cargarCatalogo, () => cargarMisPedidos(true), revisarPedidosAsignados, revisarAlertas, avisarConectado, revisarAsignadosTomadosPorOtro];
        for (const f of tareas) { try { f(); } catch (e) {} }
      }
      poner(ok ? '\u2714 Conectado' : '\u2716 Sin conexi\u00f3n con la PC', true);
      setTimeout(() => { poner('\u21bb Reconectar', false); reconectando = false; }, 2000);
    }
    function abrirMenu() { document.getElementById('menuOverlay').style.display = 'block'; }
    function activarModoCliente() { cerrarMenu(); document.body.classList.add('modoCliente'); document.getElementById('buscador').value=''; buscar(); document.getElementById('buscador').focus(); }
    // Salir del modo cliente = el vendedor vuelve a usar su telefono: se pide el
    // PIN para verificar su identidad. Se valida contra la PC (el PIN puesto
    // desde la PC tambien cuenta); si no hay conexion, se usa el PIN guardado
    // en este telefono.
    async function salirModoCliente() {
      let tienePin = false, pinLocal = miPin();
      const nombre = (nombreInput.value || '').trim();
      let servidorOk = false;
      try {
        const r = await fetch('/api/pines/verificar', { method: 'POST', body: JSON.stringify({ vendedor: nombre, pin: '' }) });
        const d = await r.json();
        tienePin = !!d.tienePin; servidorOk = true;
      } catch (e) { tienePin = !!pinLocal; }
      if (tienePin) {
        const intento = prompt('Escribe tu PIN para volver al modo vendedor:');
        if (intento === null) return;
        let valido = false;
        if (servidorOk) {
          try {
            const r = await fetch('/api/pines/verificar', { method: 'POST', body: JSON.stringify({ vendedor: nombre, pin: intento.trim() }) });
            valido = !!(await r.json()).valido;
          } catch (e) { valido = intento.trim() === pinLocal; }
        } else {
          valido = intento.trim() === pinLocal;
        }
        if (!valido) { alert('PIN incorrecto.'); return; }
      }
      document.body.classList.remove('modoCliente');
      abrirMenu();
    }
    function cerrarMenu() { document.getElementById('menuOverlay').style.display = 'none'; pintarVendedorHeader(); restaurarMetodoPagoPreferido(); }
    pintarVendedorHeader();

    // ---- Inicio de sesion del vendedor ----
    // El vendedor SOLO existe si la PC lo creo (con su PIN). Nadie puede
    // entrar con un nombre que no este registrado. Cada telefono guarda el
    // PIN ya verificado, asi que su dueno solo lo escribe la primera vez (o
    // si el PIN cambia). Sin conexion con la PC, solo se puede reentrar con
    // un nombre que este mismo telefono ya verifico antes contra el servidor
    // (guardado en 'vendedorValidado'); un nombre nuevo nunca entra sin red.
    async function consultarPinVendedor(nombre, pin) {
      const r = await fetch('/api/pines/verificar', { method: 'POST', body: JSON.stringify({ vendedor: nombre, pin: pin }) });
      return await r.json();
    }
    async function iniciarSesionVendedor() {
      if (autoservicioActivo) return true;
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) { abrirLogin(); return false; }
      let d;
      try { d = await consultarPinVendedor(nombre, miPin()); } catch (e) { return true; }
      if (!d.existe) { abrirLogin(); return false; }
      loginPinLargo = d.largoPin || 6;
      if (d.valido) return true;
      abrirLogin(nombre);
      return false;
    }

    // ---- Pantalla de login: usuario (de una lista de registrados), luego PIN con teclado numerico ----
    let loginNombre = '', loginPin = '', loginOcupado = false, loginPinLargo = 6;
    const loginEl = document.getElementById('loginOverlay');
    function loginVisible() { return loginEl.style.display !== 'none'; }
    function setLoginError(t) { document.getElementById('loginError').textContent = t || ''; }
    async function cargarListaUsuariosLogin() {
      try {
        const res = await fetch('/api/pines');
        const data = await res.json();
        const lista = (data && data.vendedoresConPin) || [];
        document.getElementById('dlUsuariosLogin').innerHTML = lista.map(v => '<option value="' + v + '">').join('');
      } catch (e) {}
    }
    function abrirLogin(nombre) {
      loginPin = '';
      loginOcupado = false;
      loginEl.style.display = 'flex';
      document.getElementById('loginOlvido').style.display = 'none';
      if (autoservicioActivo) { prepararLoginCliente(); mostrarPasoNombre(); return; }
      cargarListaUsuariosLogin();
      if (nombre) { loginNombre = nombre; mostrarPasoPin(); } else { mostrarPasoNombre(); }
    }
    function cerrarLogin() { loginEl.style.display = 'none'; }
    function mostrarPasoNombre() {
      loginNombre = ''; loginPin = '';
      document.getElementById('loginPasoNombre').style.display = 'block';
      document.getElementById('loginPasoPin').style.display = 'none';
      const inp = document.getElementById('loginNombreInput');
      inp.value = '';
      setLoginError('');
      setTimeout(() => inp.focus(), 60);
    }
    function mostrarPasoPin() {
      document.getElementById('loginPasoNombre').style.display = 'none';
      document.getElementById('loginPasoPin').style.display = 'block';
      document.getElementById('loginNombreTxt').textContent = loginNombre;
      setLoginError('');
      const puntos = document.getElementById('loginPuntos');
      puntos.innerHTML = '';
      for (let i = 0; i < loginPinLargo; i++) puntos.appendChild(document.createElement('span'));
      pintarPuntosPin();
    }
    function pintarPuntosPin() {
      document.querySelectorAll('#loginPuntos span').forEach((p, i) => p.classList.toggle('lleno', i < loginPin.length));
      document.getElementById('btnLoginEntrar').disabled = loginOcupado || loginPin.length < loginPinLargo;
    }
    function loginTecla(t) {
      if (loginOcupado) return;
      setLoginError('');
      if (t === 'C') loginPin = '';
      else if (t === 'B') loginPin = loginPin.slice(0, -1);
      else if (loginPin.length < loginPinLargo) loginPin += t;
      pintarPuntosPin();
      if (loginPin.length === loginPinLargo) loginEntrar();
    }
    function entrarComoVendedor(nombre, pin) {
      nombreInput.value = nombre;
      localStorage.setItem('vendedorNombre', nombre);
      localStorage.setItem('vendedorValidado', nombre.trim().toLowerCase());
      pinInput.value = pin || '';
      localStorage.setItem('vendedorPin', pin || '');
      pinVendedorAlEnfocar = pin || '';
      cerrarLogin();
      pintarVendedorHeader();
      restaurarMetodoPagoPreferido();
      try { posAplicarUI(true); } catch (e) {}
    }
    // Paso 1: el nombre (elegido de la lista de vendedores que creo la PC, o
    // escrito a mano). Si no esta registrado, no se puede entrar.
    // Cliente (autoservicio): la misma pantalla pide SOLO el nombre. No se
    // consulta a la PC ni se pide PIN; el nombre queda guardado en este telefono.
    let enviarTrasNombreCliente = false;
    function prepararLoginCliente() {
      document.querySelector('#loginTarjeta .login-titulo').textContent = 'Tu nombre';
      document.querySelector('#loginPasoNombre .login-label').textContent = 'Nombre';
      const inp = document.getElementById('loginNombreInput');
      inp.removeAttribute('list');
      inp.placeholder = 'Escribe tu nombre';
      document.querySelector('#loginPasoNombre .login-ayuda').textContent = 'Escribe tu nombre para tu pedido';
    }
    function entrarComoCliente(nombre) {
      nombreInput.value = nombre.slice(0, 30);
      localStorage.setItem('clienteNombre', nombreInput.value);
      cerrarLogin();
      pintarVendedorHeader();
      if (enviarTrasNombreCliente) { enviarTrasNombreCliente = false; enviarPedido(); }
    }
    async function loginContinuar() {
      const nombre = document.getElementById('loginNombreInput').value.trim();
      if (!nombre) { setLoginError('Escribe tu nombre.'); return; }
      setLoginError('');
      if (autoservicioActivo) { entrarComoCliente(nombre); return; }
      let d = null;
      try { d = await consultarPinVendedor(nombre, ''); } catch (e) {}
      if (!d) {
        // Sin conexion con la PC: solo se puede reentrar con un nombre que
        // este telefono ya verifico antes (no se puede validar uno nuevo).
        const validado = (localStorage.getItem('vendedorValidado') || '').trim().toLowerCase();
        if (validado === nombre.toLowerCase() && miPin()) {
          entrarComoVendedor(nombre, miPin());
        } else {
          setLoginError('Sin conexion con la PC: no se puede verificar ese usuario todavia.');
        }
        return;
      }
      if (!d.existe) {
        setLoginError('Ese usuario no esta registrado. Pidele al encargado que te cree el usuario desde la PC.');
        return;
      }
      loginPinLargo = d.largoPin || 6;
      loginNombre = nombre; loginPin = '';
      mostrarPasoPin();
    }
    // Paso 2: el PIN.
    async function loginEntrar() {
      if (loginOcupado || loginPin.length < loginPinLargo) return;
      loginOcupado = true; pintarPuntosPin();
      let d = null;
      try { d = await consultarPinVendedor(loginNombre, loginPin); } catch (e) {}
      loginOcupado = false;
      if (!d) {
        const validado = (localStorage.getItem('vendedorValidado') || '').trim().toLowerCase();
        if (validado === loginNombre.toLowerCase() && miPin() && miPin() === loginPin) { entrarComoVendedor(loginNombre, loginPin); return; }
        loginPin = ''; pintarPuntosPin();
        setLoginError('Sin conexion con la PC. Revisa la Wi-Fi e intenta otra vez.');
        return;
      }
      if (d.valido) { entrarComoVendedor(loginNombre, loginPin); return; }
      const tarjeta = document.getElementById('loginTarjeta');
      tarjeta.classList.remove('sacudir'); void tarjeta.offsetWidth; tarjeta.classList.add('sacudir');
      loginPin = ''; pintarPuntosPin();
      setLoginError('PIN incorrecto. Intenta de nuevo.');
    }
    document.addEventListener('keydown', (e) => {
      if (!loginVisible() || document.getElementById('loginPasoPin').style.display === 'none') return;
      if (/^[0-9]$/.test(e.key)) loginTecla(e.key);
      else if (e.key === 'Backspace') loginTecla('B');
      else if (e.key === 'Enter') loginEntrar();
    });
    function cerrarSesion() {
      nombreInput.value = ''; localStorage.setItem('vendedorNombre', '');
      pinInput.value = ''; localStorage.setItem('vendedorPin', '');
      pinVendedorAlEnfocar = '';
      try { posAplicarUI(true); } catch (e) {}
      cerrarMenu();
      abrirLogin();
    }
    nombreInput.addEventListener('change', iniciarSesionVendedor);
    iniciarSesionVendedor();

    if (autoservicioActivo) {
      pintarVendedorHeader();
      if (!(nombreInput.value || '').trim()) abrirLogin();
    }
    restaurarMetodoPagoPreferido();

    // Recuerda, por vendedor (nombre escrito en Ajustes), el ultimo metodo
    // de pago que uso, para no tener que reelegirlo en cada pedido.
    function guardarMetodoPagoPreferido() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) return;
      localStorage.setItem('metodoPagoPreferido_' + nombre, document.getElementById('metodoPago').value);
    }
    function restaurarMetodoPagoPreferido() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) return;
      const guardado = localStorage.getItem('metodoPagoPreferido_' + nombre);
      const sel = document.getElementById('metodoPago');
      if (guardado && Array.from(sel.options).some(o => o.value === guardado) && sel.value !== guardado) {
        sel.value = guardado;
        renderCarrito();
      }
    }

    // ---- Catalogo (centralizado en la PC) ----
    async function cargarCatalogo() {
      const info = document.getElementById('estadoCatalogo');
      try {
        const res = await fetch('/api/catalogo');
        const datosCat = await res.json();
        if (!res.ok || !Array.isArray(datosCat)) throw new Error('catalogo');
        catalogo = datosCat;
        try {
          const resCfg = await fetch('/api/config');
          if (resCfg.ok) {
            const cfg = await resCfg.json();
            ocultarSinStock = !!cfg.ocultarSinStock;
            const tasaNueva = parseFloat(cfg.tasaDolar) || 0;
            if (tasaNueva !== tasaDolarActual) { tasaDolarActual = tasaNueva; renderCarrito(); }
            umbralStockBajoActual = parseInt(cfg.umbralStockBajo, 10) || 0;
            const descNuevo = !!cfg.permitirDescuentos;
            if (descNuevo !== permitirDescuentosActual) { permitirDescuentosActual = descNuevo; renderCarrito(); }
            autoservicioDestino = cfg.autoservicioDestino === 'vendedor' ? 'vendedor' : 'pc';
          }
        } catch(e) {}
        guardarCatalogoLocal();
        aplicarColaAlStock();
        info.classList.remove('error');
        document.getElementById('btnMenu').classList.remove('alerta');
        if (ocultandoSinStock()) {
          const conStock = catalogo.filter(p => !productoSinStock(p)).length;
          info.textContent = conStock + ' productos con stock (se ocultan ' + (catalogo.length - conStock) + ' sin stock).';
        } else {
          info.textContent = catalogo.length + ' productos disponibles.';
        }
        buscar();
      } catch(e) {
        info.classList.add('error');
        document.getElementById('btnMenu').classList.add('alerta');   // punto rojo: hay un problema en Ajustes
        if (autoservicioActivo) { info.textContent = 'No se pudo cargar el catalogo. Revisa la conexion con la PC.'; return; }
        // MODO SIN PC: se usa el ultimo catalogo guardado en este telefono
        const guardado = leerCatalogoLocal();
        if (guardado) {
          catalogo = guardado.catalogo;
          ocultarSinStock = !!guardado.ocultarSinStock;
          tasaDolarActual = guardado.tasaDolarActual || 0;
          umbralStockBajoActual = guardado.umbralStockBajoActual || 0;
          permitirDescuentosActual = !!guardado.permitirDescuentosActual;
          aplicarColaAlStock();
          const f = new Date(guardado.t);
          info.textContent = 'SIN PC: usando el catalogo guardado el ' + f.toLocaleDateString() + ' ' + f.toLocaleTimeString().slice(0, 5) + ' (' + catalogo.length + ' productos). Los pedidos se guardan y se envian solos al volver.';
          buscar();
        } else {
          info.textContent = 'No se pudo cargar el catalogo y aun no hay uno guardado en este telefono. Conectate a la PC una vez.';
        }
      }
    }
    // Copia del catalogo en el telefono (para vender sin la PC) + descuento del stock
    // de los pedidos guardados que todavia no se han enviado (asi no se vende de mas).
    const CLAVE_CATALOGO_LOCAL = 'catalogoLocalV1';
    function guardarCatalogoLocal() {
      try {
        localStorage.setItem(CLAVE_CATALOGO_LOCAL, JSON.stringify({ t: Date.now(), catalogo: catalogo, ocultarSinStock: ocultarSinStock,
          tasaDolarActual: tasaDolarActual, umbralStockBajoActual: umbralStockBajoActual, permitirDescuentosActual: permitirDescuentosActual }));
      } catch (e) {}
    }
    function leerCatalogoLocal() {
      try {
        const g = JSON.parse(localStorage.getItem(CLAVE_CATALOGO_LOCAL) || 'null');
        return (g && Array.isArray(g.catalogo) && g.catalogo.length) ? g : null;
      } catch (e) { return null; }
    }
    function aplicarColaAlStock() {
      try {
        const usado = {};
        leerCola().forEach(c => ((c.payload && c.payload.items) || []).forEach(it => { usado[it.sku] = (usado[it.sku] || 0) + Number(it.cantidad || 0); }));
        catalogo.forEach(p => { if (usado[p.sku] && p.stock !== null && p.stock !== undefined) p.stock = Math.max(0, p.stock - usado[p.sku]); });
      } catch (e) {}
    }
    cargarCatalogo();
    setInterval(cargarCatalogo, 15000);

    // ---- Subir y mapear Excel desde el movil ----
    let archivoExcelBuffer = null;
    let configMapeoEncabezados = [];

    document.getElementById('archivoExcelMovil').addEventListener('change', onArchivoExcelSeleccionado);
    document.getElementById('zipFotosMovil').addEventListener('change', onZipFotosSeleccionado);

    async function onZipFotosSeleccionado(ev) {
      const file = ev.target.files[0];
      if (!file) return;
      const estado = document.getElementById('estadoSubidaFotos');
      estado.textContent = 'Leyendo y guardando en este telefono...';
      try {
        const info = await procesarZipFotosLocal(file);
        estado.textContent = 'Listo: ' + info.copiadas + ' de ' + info.totalProductos + ' productos con foto guardados en este telefono' + (info.sinFoto ? (', ' + info.sinFoto + ' sin foto en el catalogo') : '') + '.';
        buscar();
      } catch (e) {
        estado.textContent = (e && e.message) || 'No se pudo procesar el archivo.';
      } finally {
        ev.target.value = '';
      }
    }

    // ---- Descargar el Excel que tiene configurado la PC ----
    async function descargarExcelCatalogo() {
      const estado = document.getElementById('estadoDescargaExcel');
      estado.textContent = 'Descargando...';
      try {
        const res = await fetch('/api/catalogo/archivo');
        if (!res.ok) {
          const texto = await res.text().catch(() => '');
          estado.textContent = texto || 'No se pudo descargar el Excel.';
          return;
        }
        const blob = await res.blob();
        const nombre = res.headers.get('X-Nombre-Archivo') || 'catalogo.xlsx';
        const como = await guardarArchivoEnMovil(blob, nombre);
        estado.textContent = como === 'compartido'
          ? 'Listo: elige donde guardar o con que abrir "' + nombre + '".'
          : 'Descargado: ' + nombre;
      } catch (e) {
        if (e && e.message === 'sin-soporte') {
          estado.textContent = 'Este telefono no deja guardar el archivo desde la app. Abre ' + location.origin + '/vendedor en Chrome y descargalo ahi (o actualiza la app).';
        } else if (e && (e.name === 'AbortError' || /cancel/i.test(String(e.message || '')))) {
          estado.textContent = 'Cancelado.';
        } else {
          estado.textContent = 'No se pudo descargar (' + ((e && e.message) || 'revisa la conexion con la PC') + ').';
        }
      }
    }

    // En la app instalada (APK) el WebView NO descarga archivos con <a download>: por eso antes
    // "no pasaba nada". Aqui se guarda el archivo con el plugin nativo y se abre el menu de
    // compartir de Android (guardar en Archivos / Drive, abrir con Excel, enviar por WhatsApp).
    // En Chrome / PWA se usa la descarga normal.
    function blobABase64(blob) {
      return new Promise((resolve, reject) => {
        const fr = new FileReader();
        fr.onload = () => resolve(String(fr.result).split(',')[1] || '');
        fr.onerror = () => reject(fr.error || new Error('no se pudo leer el archivo'));
        fr.readAsDataURL(blob);
      });
    }
    async function guardarArchivoEnMovil(blob, nombre) {
      const cap = window.Capacitor;
      const plug = (cap && cap.Plugins) || {};
      const enApp = !!(cap && (typeof cap.isNativePlatform === 'function' ? cap.isNativePlatform() : cap.platform === 'android'));
      if (plug.Filesystem && plug.Share) {
        const datos = await blobABase64(blob);
        const r = await plug.Filesystem.writeFile({ path: nombre, data: datos, directory: 'CACHE', recursive: true });
        await plug.Share.share({ title: nombre, url: r.uri, dialogTitle: 'Guardar o abrir el Excel' });
        return 'compartido';
      }
      const archivo = new File([blob], nombre, { type: blob.type || 'application/octet-stream' });
      if (navigator.canShare && navigator.canShare({ files: [archivo] })) {
        await navigator.share({ files: [archivo], title: nombre });
        return 'compartido';
      }
      if (enApp) throw new Error('sin-soporte');   // APK vieja sin los plugins: se avisa en vez de no hacer nada
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = nombre;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 8000);
      return 'descargado';
    }

    async function onArchivoExcelSeleccionado(ev) {
      const file = ev.target.files[0];
      if (!file) return;
      const estado = document.getElementById('estadoSubidaExcel');
      estado.textContent = 'Leyendo archivo...';
      try {
        archivoExcelBuffer = await file.arrayBuffer();
        const wb = XLSX.read(archivoExcelBuffer, { type: 'array' });
        const hoja = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });
        if (!filas.length) { estado.textContent = 'El archivo esta vacio.'; return; }

        // Si ya hay un mapeo guardado (de la PC o de otro vendedor) y coincide
        // con los encabezados de este archivo, se usa directo sin preguntar.
        const resCfg = await fetch('/api/catalogo/config');
        const cfg = await resCfg.json();
        const mapeoPrevio = cfg && cfg.mapeo;
        filaEncabezado = filaEncabezadoPara(filas, mapeoPrevio);
        configMapeoEncabezados = nombresColumnas(filas[filaEncabezado]);
        configMapeoFilasDatos = filas.slice(filaEncabezado + 1, filaEncabezado + 61);
        try { filaBaseExcel = XLSX.utils.decode_range(hoja['!ref']).s.r; } catch (e) { filaBaseExcel = 0; }
        const coincide = mapeoPrevio && mapeoPrevio.nombre && mapeoPrevio.precio &&
          configMapeoEncabezados.includes(mapeoPrevio.nombre) &&
          configMapeoEncabezados.includes(mapeoPrevio.precio) &&
          (!mapeoPrevio.sku || configMapeoEncabezados.includes(mapeoPrevio.sku)) &&
          (!mapeoPrevio.stock || configMapeoEncabezados.includes(mapeoPrevio.stock));

        if (coincide) {
          estado.textContent = 'Procesando con las columnas ya configuradas...';
          await procesarYSubirExcel(mapeoPrevio);
          estado.textContent = 'Catalogo actualizado desde tu Excel.';
          cargarCatalogo();
        } else {
          abrirMapeoManual();
        }
      } catch (e) {
        estado.textContent = 'No se pudo leer el archivo: ' + e.message;
      } finally {
        ev.target.value = '';
      }
    }

    // ---- Fila de encabezados y columnas del Excel ----
    // Muchos Excel traen filas de titulo arriba (nombre del negocio, fecha,
    // logo...) y los encabezados reales unas filas mas abajo. En vez de
    // asumir que estan en la fila 1, se busca la fila que mas parece un
    // encabezado. Las columnas sin encabezado se llaman "Columna N" y la
    // ventana de columnas muestra un ejemplo de cada una para elegir facil.
    let filaEncabezado = 0;
    let filaBaseExcel = 0;   // fila real de Excel donde empieza el rango usado (algunas hojas empiezan en B2, etc.)
    let configMapeoFilasDatos = [];

    function textoCelda(v) { return String(v == null ? '' : v); }

    function escaparHtml(t) {
      return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // Nombres de columna de una fila de encabezados: sin encabezado ->
    // "Columna N"; si dos se llaman igual, la segunda lleva " (2)".
    function nombresColumnas(fila) {
      const usados = new Map();
      return (fila || []).map((c, i) => {
        const n = textoCelda(c).trim() === '' ? ('Columna ' + (i + 1)) : textoCelda(c);
        const veces = (usados.get(n) || 0) + 1;
        usados.set(n, veces);
        return veces > 1 ? (n + ' (' + veces + ')') : n;
      });
    }

    function detectarFilaEncabezado(filas) {
      const clave = /sku|codigo|código|code|nombre|name|producto|descrip|articulo|artículo|precio|price|stock|cantidad|quantity|existencia|disponib|inventario|unidad/;
      let mejor = 0, mejorPuntos = -1;
      const limite = Math.min(filas.length, 40);
      for (let r = 0; r < limite; r++) {
        let textos = 0, claves = 0;
        for (const c of (filas[r] || [])) {
          const t = textoCelda(c).trim();
          if (!t) continue;
          if (isNaN(Number(t.replace(',', '.')))) textos++;
          if (clave.test(t.toLowerCase())) claves++;
        }
        const puntos = claves * 10 + textos;
        if (puntos > mejorPuntos) { mejorPuntos = puntos; mejor = r; }
      }
      return mejor;
    }

    // Con un mapeo ya guardado: la fila que contiene las columnas guardadas
    // de Nombre y Precio. Si no aparece (otro archivo), se detecta sola.
    function filaEncabezadoPara(filas, mapeo) {
      if (mapeo && mapeo.nombre && mapeo.precio) {
        const limite = Math.min(filas.length, 60);
        for (let r = 0; r < limite; r++) {
          const celdas = nombresColumnas(filas[r]);
          if (celdas.includes(mapeo.nombre) && celdas.includes(mapeo.precio)) return r;
        }
      }
      return detectarFilaEncabezado(filas);
    }

    // Primer valor no vacio de la columna (ejemplo para la lista)
    function ejemploColumna(i) {
      for (const f of configMapeoFilasDatos) {
        const v = textoCelda(f[i]).trim();
        if (v !== '') return v.length > 26 ? (v.slice(0, 26) + '…') : v;
      }
      return '';
    }

    function estadisticasColumna(i) {
      const vistos = new Set();
      let total = 0, textos = 0, enteros = 0;
      for (const f of configMapeoFilasDatos) {
        const t = textoCelda(f[i]).trim();
        if (t === '') continue;
        total++;
        vistos.add(t);
        if (isNaN(Number(t.replace(',', '.')))) textos++;
        else if (/^\d+$/.test(t)) enteros++;
      }
      return { total, textos, enteros, unicos: vistos.size };
    }

    // Cuando el encabezado no dice nada (columna sin titulo), se adivina por
    // lo que contiene: 'nombre' = texto casi siempre distinto; 'sku' = codigos
    // enteros casi todos distintos.
    function adivinarPorContenido(tipo, yaUsadas) {
      let mejor = -1, mejorPuntos = 0;
      for (let i = 0; i < configMapeoEncabezados.length; i++) {
        if (yaUsadas.includes(i)) continue;
        const e = estadisticasColumna(i);
        if (e.total < 3) continue;
        let puntos = 0;
        if (tipo === 'nombre') puntos = (e.textos / e.total > 0.8) ? e.unicos : 0;
        else puntos = (e.enteros / e.total > 0.9 && e.unicos / e.total > 0.9) ? e.unicos : 0;
        if (puntos > mejorPuntos) { mejorPuntos = puntos; mejor = i; }
      }
      return mejor;
    }

    // Prueba varias expresiones en orden de prioridad (la primera que
    // encuentre alguna columna gana), saltando las columnas "excluir".
    function adivinarColumna(regexes, excluir) {
      const lista = Array.isArray(regexes) ? regexes : [regexes];
      for (const rx of lista) {
        for (let i = 0; i < configMapeoEncabezados.length; i++) {
          const h = configMapeoEncabezados[i].toLowerCase().trim();
          if (!h || /^columna \d+$/.test(h)) continue;
          if (excluir && excluir.test(h)) continue;
          if (rx.test(h)) return i;
        }
      }
      return -1;
    }

    function preseleccionarColumnas() {
      const p = {
        precio: adivinarColumna(
          [/^\s*(precio\s*(de\s*)?venta|precio|pvp|retail\s*price|price)\s*$/, /precio\s*(de\s*)?venta|retail|pvp/, /precio|price/],
          /group|grupo|purchase|compra|wholesale|mayor|cost|\b[2-9]\b/),
        stock: adivinarColumna(
          [/^\s*(stock|cantidad|existencias?|disponible|final|available\s*quantity)\s*$/, /stock|cantidad|existencia|disponib|quantity|inventario/],
          /m[ií]n|nominal|group|grupo/),
        nombre: adivinarColumna(
          [/^\s*(nombre|name|producto|descripci[oó]n|art[ií]culo)\s*$/, /nombre|descrip|art[ií]culo|producto|\bname\b/],
          /c[oó]digo|code|sku|\bcod\b|\bref|group|grupo|catalog|\b[2-9]\b/),
        sku: adivinarColumna(
          [/^\s*(sku|c[oó]digo|code|cod|ref(erencia)?)\s*$/, /sku|c[oó]digo|\bcode\b|\bcod\b|\bref/],
          /barcode|bar\s*code|barras|catalog|group|grupo/)
      };
      const usadas = () => [p.precio, p.stock, p.nombre, p.sku].filter(i => i >= 0);
      if (p.nombre < 0) p.nombre = adivinarPorContenido('nombre', usadas());
      if (p.sku < 0) p.sku = adivinarPorContenido('sku', usadas());
      return p;
    }

    function abrirMapeoManual() {
      document.getElementById('configOverlay').style.display = 'flex';
      document.getElementById('configPaso2Error').textContent = '';
      const pre = preseleccionarColumnas();
      const campos = [
        { id: 'sku', etiqueta: 'SKU / Codigo', preseleccion: pre.sku, opcional: true },
        { id: 'nombre', etiqueta: 'Nombre del producto', preseleccion: pre.nombre, opcional: false },
        { id: 'precio', etiqueta: 'Precio', preseleccion: pre.precio, opcional: false },
        { id: 'stock', etiqueta: 'Cantidad / Stock', preseleccion: pre.stock, opcional: true }
      ];
      const cont = document.getElementById('mapeoCampos');
      cont.innerHTML = '<div style="font-size:12px; color:#64748b; margin-bottom:10px;">Encabezados detectados en la fila ' + (filaEncabezado + filaBaseExcel + 1) + ' del Excel.</div>' + campos.map(c => {
        const opciones = ['<option value="">' + (c.opcional ? '— Ninguna —' : '— elegir —') + '</option>']
          .concat(configMapeoEncabezados.map((h, i) => {
            const ej = ejemploColumna(i);
            if (!ej && /^Columna \d+$/.test(h)) return '';   // columna vacia y sin titulo
            return '<option value="' + i + '"' + (i === c.preseleccion ? ' selected' : '') + '>' + escaparHtml(h) + (ej ? '  —  ej: ' + escaparHtml(ej) : '') + '</option>';
          }))
          .join('');
        return '<div style="margin-bottom:10px;">' +
          '<label style="font-size:12px; display:block; margin-bottom:4px;">' + c.etiqueta + '</label>' +
          '<select id="mapeo-' + c.id + '">' + opciones + '</select>' +
        '</div>';
      }).join('');
    }

    function cerrarConfig() {
      document.getElementById('configOverlay').style.display = 'none';
    }

    async function guardarMapeoYCerrar() {
      const errEl = document.getElementById('configPaso2Error');
      errEl.textContent = '';
      const val = id => {
        const sel = document.getElementById('mapeo-' + id);
        const v = sel.value;
        return v === '' ? null : configMapeoEncabezados[parseInt(v, 10)];
      };
      const mapeo = { sku: val('sku'), nombre: val('nombre'), precio: val('precio'), stock: val('stock') };
      if (!mapeo.nombre || !mapeo.precio) {
        errEl.textContent = 'Elige al menos la columna de Nombre y la de Precio.';
        return;
      }
      await fetch('/api/catalogo/config', { method: 'POST', body: JSON.stringify({ mapeo }) });
      cerrarConfig();
      const estado = document.getElementById('estadoSubidaExcel');
      estado.textContent = 'Procesando...';
      await procesarYSubirExcel(mapeo);
      estado.textContent = 'Catalogo actualizado desde tu Excel.';
      cargarCatalogo();
    }

    async function procesarYSubirExcel(mapeo) {
      const wb = XLSX.read(archivoExcelBuffer, { type: 'array' });
      const hoja = wb.Sheets[wb.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '' });
      const filaEnc = filaEncabezadoPara(filas, mapeo);
      const encabezados = nombresColumnas(filas[filaEnc]);
      const idx = nombreCol => nombreCol ? encabezados.indexOf(nombreCol) : -1;
      const iSku = idx(mapeo.sku);
      const iNombre = idx(mapeo.nombre);
      const iPrecio = idx(mapeo.precio);
      const iStock = idx(mapeo.stock);

      const limpiarNumero = v => {
        const n = parseFloat(String(v == null ? '' : v).replace(',', '.').replace(/[^0-9.\-]/g, ''));
        return isNaN(n) ? 0 : n;
      };

      const productos = [];
      for (let r = filaEnc + 1; r < filas.length; r++) {
        const fila = filas[r];
        const nombre = String(iNombre >= 0 ? (fila[iNombre] ?? '') : '').trim();
        if (!nombre) continue;
        const sku = iSku >= 0 ? String(fila[iSku] ?? '').trim() : '';
        const precio = iPrecio >= 0 ? limpiarNumero(fila[iPrecio]) : 0;
        let stock = null;
        if (iStock >= 0) {
          const v = fila[iStock];
          if (v !== '' && v !== null && v !== undefined) stock = limpiarNumero(v);
        }
        productos.push({ sku, nombre, precio: Math.round(precio * 100) / 100, stock });
      }

      await fetch('/api/catalogo/importar', { method: 'POST', body: JSON.stringify({ productos }) });
    }

    // Quita acentos y pasa a minusculas: "bateria" encuentra "Batería" y viceversa.
    function normalizar(t) {
      return String(t == null ? '' : t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    }

    // ---- Fotos locales en este telefono (sin pasar por la PC) ----
    // Lector de .zip minimo en JS puro (sin librerias externas: el .zip que
    // genera la app del catalogo no usa zip64 ni contraseña, asi que con
    // leer el indice central y descomprimir con DecompressionStream alcanza).
    async function leerEntradasZip(buffer) {
      const dv = new DataView(buffer);
      const bytes = new Uint8Array(buffer);
      let eocd = -1;
      const desde = Math.max(0, bytes.length - 22 - 65557);
      for (let i = bytes.length - 22; i >= desde; i--) {
        if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
      }
      if (eocd < 0) throw new Error('No es un .zip valido (no se encontro el indice).');
      const totalEntradas = dv.getUint16(eocd + 10, true);
      const offCentral = dv.getUint32(eocd + 16, true);
      const entradas = [];
      let p = offCentral;
      for (let i = 0; i < totalEntradas; i++) {
        if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Indice del .zip dañado.');
        const metodo = dv.getUint16(p + 10, true);
        const compSize = dv.getUint32(p + 20, true);
        const nombreLen = dv.getUint16(p + 28, true);
        const extraLen = dv.getUint16(p + 30, true);
        const comentLen = dv.getUint16(p + 32, true);
        const offsetLocal = dv.getUint32(p + 42, true);
        const nombre = new TextDecoder('utf-8').decode(bytes.subarray(p + 46, p + 46 + nombreLen));
        entradas.push({ nombre, metodo, compSize, offsetLocal });
        p += 46 + nombreLen + extraLen + comentLen;
      }
      return entradas;
    }
    async function leerContenidoEntradaZip(buffer, entrada) {
      const dv = new DataView(buffer);
      const bytes = new Uint8Array(buffer);
      const p = entrada.offsetLocal;
      if (dv.getUint32(p, true) !== 0x04034b50) throw new Error('Cabecera local del .zip dañada.');
      const nombreLen = dv.getUint16(p + 26, true);
      const extraLen = dv.getUint16(p + 28, true);
      const inicio = p + 30 + nombreLen + extraLen;
      const datos = bytes.subarray(inicio, inicio + entrada.compSize);
      if (entrada.metodo === 0) return datos;
      if (entrada.metodo === 8) {
        const stream = new Blob([datos]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        return new Uint8Array(await new Response(stream).arrayBuffer());
      }
      throw new Error('Metodo de compresion del .zip no soportado.');
    }

    // Las fotos quedan en una base local del navegador (IndexedDB), separada
    // por telefono: no se comparten solas con la PC ni con otros vendedores.
    function abrirFotosLocalDB() {
      return new Promise((resolve, reject) => {
        const req = indexedDB.open('fotosVendedorLocal', 1);
        req.onupgradeneeded = () => { req.result.createObjectStore('fotos'); };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    }
    function guardarFotoLocal(db, sku, blob) {
      return new Promise((resolve, reject) => {
        const tx = db.transaction('fotos', 'readwrite');
        tx.objectStore('fotos').put(blob, sku);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    }
    function obtenerFotoLocal(db, sku) {
      return new Promise((resolve) => {
        try {
          const tx = db.transaction('fotos', 'readonly');
          const req = tx.objectStore('fotos').get(sku);
          req.onsuccess = () => resolve(req.result || null);
          req.onerror = () => resolve(null);
        } catch (e) { resolve(null); }
      });
    }
    let fotosLocalDBPromise = null;
    function miFotosDB() {
      if (!fotosLocalDBPromise) fotosLocalDBPromise = abrirFotosLocalDB().catch(() => null);
      return fotosLocalDBPromise;
    }
    const urlsFotoLocalCache = {}; // sku -> object URL ya generada, para no releer IndexedDB en cada busqueda

    // ---- Fotos sin parpadeo: solo se pide la imagen si el producto TIENE foto ----
    // (en la PC o guardada en este telefono). Un producto sin foto ya no hace peticiones
    // que fallan ni se "repinta" cada pocos segundos.
    let skusConFoto = null;               // SKU con foto en la PC (null = aun no se sabe)
    const skusFotoLocal = new Set();      // SKU con foto guardada en este telefono
    const fotosFallidas = new Set();      // SKU cuya foto fallo una vez: no se vuelve a pedir
    try { const g = JSON.parse(localStorage.getItem('skusConFotoPC') || 'null'); if (Array.isArray(g)) skusConFoto = new Set(g.map(String)); } catch (e) {}
    async function cargarSkusConFoto() {
      try {
        const r = await fetch('/api/fotos/skus');
        const d = await r.json();
        if (Array.isArray(d)) {
          skusConFoto = new Set(d.map(String));
          try { localStorage.setItem('skusConFotoPC', JSON.stringify(d)); } catch (e) {}
        }
      } catch (e) {}
    }
    async function cargarSkusFotoLocal() {
      try {
        const db = await miFotosDB();
        if (!db) return;
        const claves = await new Promise((resolve) => {
          try {
            const req = db.transaction('fotos', 'readonly').objectStore('fotos').getAllKeys();
            req.onsuccess = () => resolve(req.result || []);
            req.onerror = () => resolve([]);
          } catch (e) { resolve([]); }
        });
        claves.forEach(k => skusFotoLocal.add(String(k)));
      } catch (e) {}
    }
    function tieneFoto(sku) {
      if (!sku) return false;
      sku = String(sku);
      if (skusFotoLocal.has(sku)) return true;
      if (fotosFallidas.has(sku)) return false;
      return skusConFoto ? skusConFoto.has(sku) : true;
    }
    function fotoFallo(img, sku) { fotosFallidas.add(String(sku)); if (img && img.remove) img.remove(); }
    // Cambia el contenido solo si de verdad es distinto (evita repintar la lista sin necesidad).
    function ponerHtmlSiCambio(el, html) {
      const hijo = el.firstElementChild;
      if (html !== '' && hijo && hijo.__marcaHtml === html && el.__htmlPrev === html) return false;
      el.innerHTML = html;
      el.__htmlPrev = html;
      if (el.firstElementChild) el.firstElementChild.__marcaHtml = html;
      return true;
    }
    cargarSkusFotoLocal();
    cargarSkusConFoto();
    setInterval(cargarSkusConFoto, 60000);

    async function procesarZipFotosLocal(file) {
      const buffer = await file.arrayBuffer();
      const entradas = await leerEntradasZip(buffer);
      const entradaDatos = entradas.find(e => e.nombre === 'datos.json');
      if (!entradaDatos) throw new Error('El .zip no tiene "datos.json" (¿es una copia de seguridad valida del catalogo?).');
      const datos = JSON.parse(new TextDecoder('utf-8').decode(await leerContenidoEntradaZip(buffer, entradaDatos)));
      const productos = datos.products || [];
      const db = await abrirFotosLocalDB();
      let copiadas = 0, sinFoto = 0;
      for (const p of productos) {
        const sku = (p.sku || '').toString().trim();
        if (!sku) continue;
        const entradaFoto = entradas.find(e => e.nombre === ('fotos/' + p.id + '.jpg'));
        if (!entradaFoto) { sinFoto++; continue; }
        const blob = new Blob([await leerContenidoEntradaZip(buffer, entradaFoto)], { type: 'image/jpeg' });
        await guardarFotoLocal(db, sku, blob);
        delete urlsFotoLocalCache[sku];
        copiadas++;
      }
      return { copiadas, sinFoto, totalProductos: productos.length };
    }

    // Despues de pintar los resultados, revisa cual de las miniaturas tiene
    // foto guardada localmente en este telefono y la pone ahi (sin red).
    async function pintarFotosLocales(cont) {
      const db = await miFotosDB();
      if (!db) return;
      const imgs = cont.querySelectorAll('img.miniatura[data-sku]');
      for (const img of imgs) {
        const sku = img.getAttribute('data-sku');
        if (urlsFotoLocalCache[sku]) { img.src = urlsFotoLocalCache[sku]; continue; }
        const blob = await obtenerFotoLocal(db, sku);
        if (blob) {
          const url = URL.createObjectURL(blob);
          urlsFotoLocalCache[sku] = url;
          img.src = url;
        }
      }
    }

    function buscar() {
      const q = normalizar(document.getElementById('buscador').value.trim());
      const cont = document.getElementById('resultados');
      // "Modo cliente" (el vendedor le pasa el telefono al cliente desde Ajustes):
      // catalogo completo de entrada, con el buscador para filtrar. En el
      // autoservicio por enlace/QR y en el modo vendedor normal solo salen
      // resultados de lo que se escribe (max. 25).
      const enModoCliente = document.body.classList.contains('modoCliente');
      if (!q && !enModoCliente) { cont.innerHTML = ''; renderRecientes(); return; }
      const base = ocultandoSinStock() ? catalogo.filter(p => !productoSinStock(p)) : catalogo;
      const filtrados = q
        ? base.filter(p => normalizar(p.nombre).includes(q) || normalizar(p.sku || '').includes(q))
        : base;
      const encontrados = enModoCliente ? filtrados : filtrados.slice(0, 25);

      const htmlRes = encontrados.map(p => {
        const enCarrito = carrito.find(i => i.sku === p.sku);
        const tieneStock = (p.stock !== null && p.stock !== undefined);
        const disponible = tieneStock ? (p.stock - (enCarrito ? enCarrito.cantidad : 0)) : null;
        const agotado = tieneStock && disponible <= 0;
        const stockBajo = tieneStock && !agotado && umbralStockBajoActual > 0 && disponible <= umbralStockBajoActual;
        const stockTxt = (tieneStock && !autoservicioActivo)
          ? ('<div class="sku"' + (stockBajo ? ' style="color:#f87171; font-weight:bold;"' : '') + '>' + (stockBajo ? '&#9888; ' : '') + 'Disponible: ' + disponible + '</div>')
          : '';
        const fotoTag = tieneFoto(p.sku) ? ('<img class="miniatura" data-sku="' + p.sku + '" src="/foto/' + encodeURIComponent(p.sku) + '.jpg" onerror=\'fotoFallo(this,' + JSON.stringify(String(p.sku)) + ')\' onclick=\'verFotoProducto(' + JSON.stringify(p.sku) + ',' + JSON.stringify(p.nombre) + ')\'>') : '';
        return '<div class="resultado">' +
          fotoTag +
          '<div class="info">' + p.nombre + '<div class="sku">' + (p.sku || '') + '</div><div class="precio">$' + p.precio.toFixed(2) + '</div>' + stockTxt + '</div>' +
          '<button ' + (agotado ? 'disabled' : '') + ' onclick=\'agregarAlCarrito(' + JSON.stringify(p) + ')\'>+</button>' +
        '</div>';
      }).join('');
      document.getElementById('recientes').innerHTML = '';
      if (ponerHtmlSiCambio(cont, htmlRes)) pintarFotosLocales(cont);
    }

    // ---- Agregados hace poco: acceso rapido para repetir el mismo producto ----
    let recientesSkus = [];
    function registrarReciente(sku) {
      recientesSkus = recientesSkus.filter(s => s !== sku);
      recientesSkus.unshift(sku);
      if (recientesSkus.length > 4) recientesSkus.length = 4;
    }
    function renderRecientes() {
      const cont = document.getElementById('recientes');
      if (!cont) return;
      if ((document.getElementById('buscador').value || '').trim() || recientesSkus.length === 0) { cont.innerHTML = ''; return; }
      const items = recientesSkus.map(sku => catalogo.find(p => p.sku === sku)).filter(Boolean);
      if (items.length === 0) { cont.innerHTML = ''; return; }
      cont.innerHTML = '<div style="width:100%; font-size:12px; color:#94a3b8; margin-bottom:2px;">Agregados hace poco (toca para repetir):</div>' +
        items.map(p => '<button class="chip-reciente" onclick=\'agregarAlCarrito(' + JSON.stringify(p) + ')\'>' + p.nombre + ' · $' + p.precio.toFixed(2) + '</button>').join('');
    }

    // ---- Foto ampliada del producto: primero mira si esta guardada en este
    // telefono (local), y si no, la pide a la PC (/foto/grande/<sku>.jpg) ----
    async function verFotoProducto(sku, nombre) {
      document.getElementById('fotoOverlayNombre').textContent = nombre;
      document.getElementById('fotoOverlay').style.display = 'flex';
      const img = document.getElementById('fotoOverlayImg');
      if (urlsFotoLocalCache[sku]) { img.src = urlsFotoLocalCache[sku]; return; }
      const db = await miFotosDB();
      const blob = db ? await obtenerFotoLocal(db, sku) : null;
      if (blob) {
        const url = URL.createObjectURL(blob);
        urlsFotoLocalCache[sku] = url;
        img.src = url;
      } else {
        img.src = '/foto/grande/' + encodeURIComponent(sku) + '.jpg';
      }
    }
    function cerrarFotoProducto() {
      document.getElementById('fotoOverlay').style.display = 'none';
    }

    // ---- Carrito ----
    function agregarAlCarrito(p) {
      const existente = carrito.find(i => i.sku === p.sku);
      const enCarritoActual = existente ? existente.cantidad : 0;
      if (p.stock !== null && p.stock !== undefined && (enCarritoActual + 1) > p.stock) {
        alert(autoservicioActivo ? ('No hay mas unidades disponibles de "' + p.nombre + '".') : ('Solo quedan ' + p.stock + ' disponibles de "' + p.nombre + '".'));
        return;
      }
      if (carrito.length === 0) {
        document.getElementById('accionesPostEnvio').style.display = 'none';
      }
      if (existente) existente.cantidad++;
      else carrito.push({ sku:p.sku, nombre:p.nombre, precio:p.precio, precioLista:p.precio, cantidad:1, stock:(p.stock === undefined ? null : p.stock) });
      renderCarrito();
      registrarReciente(p.sku);
      // Al agregar, se limpia y se encoge el buscador: queda listo para el siguiente producto.
      document.getElementById('buscador').value = '';
      document.getElementById('resultados').innerHTML = '';
      renderRecientes();
    }

    function cambiarCantidad(sku, delta) {
      const item = carrito.find(i => i.sku === sku);
      if (!item) return;
      if (delta > 0 && item.stock !== null && item.stock !== undefined && (item.cantidad + 1) > item.stock) {
        alert(autoservicioActivo ? ('No hay mas unidades disponibles de "' + item.nombre + '".') : ('Solo quedan ' + item.stock + ' disponibles de "' + item.nombre + '".'));
        return;
      }
      item.cantidad += delta;
      if (item.cantidad <= 0) carrito = carrito.filter(i => i.sku !== sku);
      renderCarrito();
      buscar();
    }

    // Escribir la cantidad directamente (toca el numero y teclea).
    function fijarCantidad(sku, valor) {
      const item = carrito.find(i => i.sku === sku);
      if (!item) return;
      let n = parseFloat(String(valor).replace(',', '.'));
      if (isNaN(n)) { renderCarrito(); return; }   // campo vacio: vuelve al valor anterior
      n = Math.round(n * 100) / 100;
      if (item.stock !== null && item.stock !== undefined && n > item.stock) {
        alert(autoservicioActivo ? ('No hay mas unidades disponibles de "' + item.nombre + '".') : ('Solo quedan ' + item.stock + ' disponibles de "' + item.nombre + '".'));
        n = item.stock;
      }
      if (n <= 0) carrito = carrito.filter(i => i.sku !== sku);
      else item.cantidad = n;
      renderCarrito();
      buscar();
    }

    // Descuento por producto: el vendedor edita directamente el precio de esa
    // linea (no existe para el cliente en autoservicio). El precio del
    // catalogo no se toca; solo cambia lo que se cobra en ESTE pedido.
    function fijarPrecio(sku, valor) {
      const item = carrito.find(i => i.sku === sku);
      if (!item || !descuentosDisponibles() || autoservicioActivo) return;
      let n = parseFloat(String(valor).replace(',', '.'));
      if (isNaN(n) || n < 0) { renderCarrito(); return; }   // campo invalido: vuelve al valor anterior
      item.precio = Math.round(n * 100) / 100;
      renderCarrito();
    }

    // Si el metodo de pago elegido es Transferencia, el precio de cada
    // producto y el total se muestran ya duplicados (precio x2), en vivo,
    // igual que se cobrara y se imprimira en el recibo.
    function factorPagoActual() {
      const sel = document.getElementById('metodoPago');
      return (sel && sel.value === 'Transferencia') ? 2 : 1;
    }

    function renderCarrito() {
      const cont = document.getElementById('carrito');
      const factor = factorPagoActual();
      if (carrito.length === 0) {
        cont.innerHTML = '<div style="color:#94a3b8; font-size:13px;">Sin productos aun.</div>';
      } else {
        cont.innerHTML = carrito.map(i => {
          const precioMostrado = i.precio * factor;
          const notaTransf = factor === 2 ? ' (transferencia)' : '';
          const precioHtml = (autoservicioActivo || !descuentosDisponibles() || !puede('descuentos'))
            ? ('$' + precioMostrado.toFixed(2) + ' c/u' + notaTransf)
            : ('$<input type="number" class="qty-input" style="width:64px;" inputmode="decimal" min="0" step="any" value="' + i.precio.toFixed(2) + '" onfocus="this.select()" onchange="fijarPrecio(\'' + i.sku + '\', this.value)"> c/u' + (factor === 2 ? ' (x2 en transferencia)' : '') + ' — toca para aplicar descuento');
          return '<div class="carrito-item">' +
            '<div>' + i.nombre + '<div class="sku">' + precioHtml + '</div></div>' +
            '<div class="qty-controls">' +
              '<button onclick="cambiarCantidad(\'' + i.sku + '\', -1)">-</button>' +
              '<input type="number" class="qty-input" inputmode="decimal" min="0" step="any" value="' + i.cantidad + '" onfocus="this.select()" onchange="fijarCantidad(\'' + i.sku + '\', this.value)">' +
              '<button onclick="cambiarCantidad(\'' + i.sku + '\', 1)">+</button>' +
            '</div>' +
          '</div>';
        }).join('');
      }
      const totalBase = carrito.reduce((s,i) => s + i.precio * i.cantidad, 0);
      const total = totalBase * factor;
      document.getElementById('totalPedido').textContent = total.toFixed(2);
      document.getElementById('totalPedidoSticky').textContent = '$' + total.toFixed(2);
      document.getElementById('notaTransferenciaTotal').style.display = (factor === 2) ? 'inline' : 'none';
      const elUSD = document.getElementById('totalUSDSticky');
      if (tasaDolarActual > 0 && total > 0) {
        elUSD.textContent = '≈ $' + (total / tasaDolarActual).toFixed(2) + ' USD';
        elUSD.style.display = 'block';
      } else {
        elUSD.style.display = 'none';
      }
      calcularCambio();
      try { actualizarBotonEnviar(); } catch (e) {}
    }
    renderCarrito();

    // ---- Cobro ----
    function toggleCobro() {
      const cobrado = document.querySelector('input[name=estadoPago]:checked').value === 'cobrado';
      document.getElementById('camposCobro').style.display = cobrado ? 'block' : 'none';
      calcularCambio();
    }

    function calcularCambio() {
      const totalProductos = carrito.reduce((s,i) => s + i.precio * i.cantidad, 0);
      const estado = document.querySelector('input[name=estadoPago]:checked').value;
      const metodo = document.getElementById('metodoPago').value;
      const esTransferencia = metodo === 'Transferencia';
      const esCombinado = metodo === 'Combinado';
      const totalCobrar = esTransferencia ? totalProductos * 2 : totalProductos;

      document.getElementById('montoRecibido').style.display = esCombinado ? 'none' : 'block';
      document.getElementById('camposCobroCombinado').style.display = esCombinado ? 'block' : 'none';

      const totalTexto = document.getElementById('totalCobroTexto');
      const el = document.getElementById('cambioTexto');

      if (esCombinado) {
        const efectivoParte = parseFloat(document.getElementById('montoEfectivoCombo').value) || 0;
        const transferParte = parseFloat(document.getElementById('montoTransferCombo').value) || 0;
        const cubierto = efectivoParte + transferParte;
        const aCobrar = efectivoParte + (transferParte * 2);
        const usdTxt = (tasaDolarActual > 0 && aCobrar > 0) ? ('  (≈ $' + (aCobrar / tasaDolarActual).toFixed(2) + ' USD)') : '';
        totalTexto.textContent = 'Pedido: $' + totalProductos.toFixed(2) + ' — A cobrar (efectivo + transferencia x2): $' + aCobrar.toFixed(2) + usdTxt;
        const diff = totalProductos - cubierto;
        if (Math.abs(diff) < 0.01) {
          el.textContent = 'Las dos partes cubren el pedido completo.';
          el.style.color = '#166534';
        } else if (diff > 0) {
          el.textContent = 'Falta cubrir $' + diff.toFixed(2) + ' del pedido entre las dos partes.';
          el.style.color = '#991b1b';
        } else {
          el.textContent = 'Las partes suman $' + Math.abs(diff).toFixed(2) + ' de mas. Ajustalas para que sumen el total del pedido.';
          el.style.color = '#991b1b';
        }
        return;
      }

      // Con "Pendiente de pago" el metodo tambien cuenta: la caja lo ve y lo cobra asi.
      const usdTxt = (tasaDolarActual > 0 && totalCobrar > 0) ? ('  (≈ $' + (totalCobrar / tasaDolarActual).toFixed(2) + ' USD)') : '';
      totalTexto.textContent = (estado === 'cobrado' ? 'Total a cobrar' : 'Se cobrara en caja') +
        (esTransferencia ? ' (transferencia x2)' : '') + ': $' + totalCobrar.toFixed(2) + usdTxt;

      const montoInput = document.getElementById('montoRecibido');
      if (montoInput.value === '') { el.textContent = ''; return; }
      const recibido = parseFloat(montoInput.value) || 0;
      const cambio = recibido - totalCobrar;
      el.textContent = cambio >= 0 ? ('Cambio a devolver: $' + cambio.toFixed(2)) : ('Falta: $' + Math.abs(cambio).toFixed(2));
      el.style.color = cambio >= 0 ? '#166534' : '#991b1b';
    }

    // ---- Enviar ----
    let ultimoPedidoId = null;

    let enviandoPedido = false;   // evita que un doble toque en "Enviar pedido" duplique el pedido

    async function enviarPedido() {
      // Si el vendedor esta escribiendo una cantidad, se confirma antes de enviar.
      const activo = document.activeElement;
      if (activo && activo.classList && activo.classList.contains('qty-input')) activo.blur();
      if (carrito.length === 0) {
        const notaSola = document.getElementById('notaPedido');
        if (!autoservicioActivo && notaSola && notaSola.value.trim() && puede('mensajes')) { return enviarMensajeSolo(); }
        alert('Agrega al menos un producto.'); return;
      }
      if (!(nombreInput.value || '').trim()) {
        if (autoservicioActivo) { enviarTrasNombreCliente = true; }
        abrirLogin(); return;
      }
      const estadoChk = document.querySelector('input[name=estadoPago]:checked').value;
      if (estadoChk === 'cobrado') {
        const metodoChk = document.getElementById('metodoPago').value;
        if (metodoChk === 'Combinado') {
          const totalProductosChk = carrito.reduce((s,i) => s + i.precio * i.cantidad, 0);
          const efectivoChk = parseFloat(document.getElementById('montoEfectivoCombo').value) || 0;
          const transferChk = parseFloat(document.getElementById('montoTransferCombo').value) || 0;
          if (document.getElementById('montoEfectivoCombo').value === '' && document.getElementById('montoTransferCombo').value === '') {
            alert('Escribe cuanto se paga en efectivo y cuanto en transferencia antes de enviar.');
            document.getElementById('montoEfectivoCombo').focus();
            return;
          }
          if (Math.abs((efectivoChk + transferChk) - totalProductosChk) > 0.01) {
            alert('Las partes en efectivo y transferencia deben sumar exactamente el total del pedido ($' + totalProductosChk.toFixed(2) + '). Revisa los montos antes de enviar.');
            document.getElementById('montoEfectivoCombo').focus();
            return;
          }
        } else if (document.getElementById('montoRecibido').value === '') {
          alert('Escribe el monto recibido del cliente antes de enviar (o el cambio no va a salir bien en el recibo).');
          document.getElementById('montoRecibido').focus();
          return;
        } else {
        // No se manda el pedido si el cliente dio menos de lo que cuesta:
        // asi no se cobra de menos por un monto mal escrito.
        const totalProductosChk = carrito.reduce((s,i) => s + i.precio * i.cantidad, 0);
        const totalCobrarChk = metodoChk === 'Transferencia' ? totalProductosChk * 2 : totalProductosChk;
        const recibidoChk = parseFloat(document.getElementById('montoRecibido').value) || 0;
        if (recibidoChk < totalCobrarChk) {
          alert('El monto recibido ($' + recibidoChk.toFixed(2) + ') es menor que el total a cobrar ($' + totalCobrarChk.toFixed(2) + '). Revisa el monto antes de enviar.');
          document.getElementById('montoRecibido').focus();
          return;
        }
        }
      }
      if (enviandoPedido) return;   // ya hay un envio de este pedido en camino: se ignora el toque extra
      enviandoPedido = true;
      const btnEnviar = document.getElementById('btnEnviarPedido');
      const textoOriginalBtn = btnEnviar.textContent;
      btnEnviar.disabled = true;
      btnEnviar.textContent = 'Enviando...';

      const estado = document.querySelector('input[name=estadoPago]:checked').value;
      const metodo = document.getElementById('metodoPago').value;
      const totalProductos = carrito.reduce((s,i) => s + i.precio * i.cantidad, 0);
      const esTransferencia = metodo === 'Transferencia';
      const esCombinado = metodo === 'Combinado';
      const pagoEfectivo = esCombinado ? (parseFloat(document.getElementById('montoEfectivoCombo').value) || 0) : null;
      const pagoTransferencia = esCombinado ? (parseFloat(document.getElementById('montoTransferCombo').value) || 0) : null;
      const totalCobrar = esCombinado ? (pagoEfectivo + pagoTransferencia * 2) : (esTransferencia ? totalProductos * 2 : totalProductos);
      const recibido = esCombinado ? totalCobrar : (parseFloat(document.getElementById('montoRecibido').value) || 0);

      const payload = {
        clienteId: 'c' + Date.now() + '_' + Math.random().toString(36).slice(2, 10),
        vendedor: nombreInput.value || 'Sin nombre',
        pin: autoservicioActivo ? '' : miPin(),
        autoservicio: autoservicioActivo,
        items: carrito,
        estado: autoservicioActivo ? 'pendiente' : estado,
        metodoPago: metodo,
        montoRecibido: estado === 'cobrado' ? recibido : null,
        totalCobrado: Number(totalCobrar.toFixed(2)),
        cambio: estado === 'cobrado' ? Number((recibido - totalCobrar).toFixed(2)) : null,
        pagoEfectivo: pagoEfectivo,
        pagoTransferencia: pagoTransferencia,
        origenAsignados: origenesAsignadosCarrito,
        nota: autoservicioActivo ? '' : ((document.getElementById('notaPedido').value || '').trim())
      };

      try {
        const porAutoservicioAVendedor = autoservicioActivo && autoservicioDestino === 'vendedor';
        const urlEnvio = porAutoservicioAVendedor ? '/api/pedidos/asignar' : '/api/pedidos';
        const bodyEnvio = porAutoservicioAVendedor ? { todos: true, items: carrito, cliente: (nombreInput.value || '').trim(), claveEnvio: payload.clienteId } : payload;
        const res = await fetch(urlEnvio, { method:'POST', body: JSON.stringify(bodyEnvio) });
        const data = await res.json().catch(() => null);
        if (!res.ok) {
          const msg = (data && data.error) ? data.error : 'No se pudo enviar el pedido.';
          mostrarMensaje(msg, false);
          if (data && data.requierePin && !autoservicioActivo) abrirLogin((nombreInput.value || '').trim());
          return;
        }
        mostrarMensaje(autoservicioActivo ? 'Pedido enviado. Gracias, en un momento te atienden.' : 'Pedido enviado correctamente.', true);
        if (!porAutoservicioAVendedor && data && data.id) {
          ultimoPedidoId = data.id;
          document.getElementById('accionesPostEnvio').style.display = 'block';
          expandidosMP.add(data.id);   // se muestra desplegado para que el vendedor lo repase
        }
        carrito = [];
        origenesAsignadosCarrito = [];
        try { document.getElementById('notaPedido').value = ''; } catch (e) {}
        renderCarrito();
        document.getElementById('buscador').value = '';
        document.getElementById('resultados').innerHTML = '';
        document.getElementById('montoRecibido').value = '';
        aplicarEstadoPagoPorDefecto();
        cargarCatalogo();
        if (autoservicioActivo) { return; }
        // Se abre "Mis pedidos de hoy" con el pedido recien enviado desplegado, para
        // que el vendedor repase lo que anoto y vea si se quedo algo. Va en su propio
        // try: un fallo aqui nunca debe mandar el pedido a la cola otra vez.
        try {
          abrirMisPedidos(true);
          await cargarMisPedidos();
          const secMP = document.getElementById('seccionMisPedidos');
          if (secMP && secMP.scrollIntoView) secMP.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch (e2) {}
      } catch(e) {
        if (autoservicioActivo) {
          mostrarMensaje('Sin conexion: no se pudo enviar el pedido. Revisa tu WiFi e intenta de nuevo.', false);
        } else {
          agregarACola(payload);
          try { document.getElementById('notaPedido').value = ''; } catch (e) {}
          mostrarMensaje('Sin conexion: el pedido se guardo en el telefono y se enviara solo cuando vuelva la red.', true);
        }
        carrito = [];
        renderCarrito();
      } finally {
        enviandoPedido = false;
        btnEnviar.disabled = false;
        btnEnviar.textContent = textoOriginalBtn;
        try { actualizarBotonEnviar(); } catch (e) {}
        document.getElementById('buscador').value = '';
        document.getElementById('resultados').innerHTML = '';
        document.getElementById('montoRecibido').value = '';
        aplicarEstadoPagoPorDefecto();
        if (autoservicioActivo) buscar();
      }
    }

    async function imprimirUltimoPedido() {
      if (!ultimoPedidoId) return;
      try {
        const res = await fetch('/api/pedidos/' + ultimoPedidoId + '/imprimir', { method: 'POST' });
        const data = await res.json().catch(() => null);
        if (data && data.ok) {
          mostrarMensaje('Recibo enviado a imprimir.', true);
        } else {
          mostrarMensaje('No se pudo imprimir: ' + ((data && data.error) || 'revisa la impresora en la PC'), false);
        }
      } catch (e) {
        mostrarMensaje('No se pudo imprimir. Revisa la conexion con la PC.', false);
      }
    }

    function mostrarMensaje(texto, ok) {
      const m = document.getElementById('mensaje');
      m.textContent = texto;
      m.className = ok ? 'ok' : 'error';
      m.style.display = 'block';
      setTimeout(() => { m.style.display = 'none'; }, 4000);
    }

    // ---- Cola offline: pedidos que no se pudieron enviar por falta de red ----
    const CLAVE_COLA = 'colaPedidosOffline';

    function leerCola() {
      try { return JSON.parse(localStorage.getItem(CLAVE_COLA) || '[]'); } catch(e) { return []; }
    }
    function guardarCola(cola) {
      localStorage.setItem(CLAVE_COLA, JSON.stringify(cola));
      actualizarColaUI();
    }
    function agregarACola(payload) {
      const cola = leerCola();
      cola.push({ id: 'off_' + Date.now() + '_' + Math.floor(Math.random() * 1000), payload, error: null });
      guardarCola(cola);
    }
    function descartarDeCola(id) {
      if (!confirm('¿Descartar este pedido guardado? No se enviara ni se contara como venta.')) return;
      guardarCola(leerCola().filter(c => c.id !== id));
    }

    function actualizarColaUI() {
      const cola = leerCola();
      const cont = document.getElementById('colaOfflineAviso');
      if (cola.length === 0) { cont.style.display = 'none'; cont.innerHTML = ''; return; }
      cont.style.display = 'block';
      cont.innerHTML = cola.map(c => {
        const totalTxt = '$' + c.payload.items.reduce((s, it) => s + it.precio * it.cantidad, 0).toFixed(2);
        if (c.error) {
          return '<div style="margin-bottom:6px;">⚠ Pedido sin enviar (' + totalTxt + '): ' + c.error +
            ' <button onclick="intentarEnviarCola()">Reintentar</button> ' +
            '<button onclick="descartarDeCola(\'' + c.id + '\')">Descartar</button></div>';
        }
        return '<div style="margin-bottom:6px;">⏳ Pedido guardado (' + totalTxt + '), esperando conexion...</div>';
      }).join('') + (cola.length > 1 ? '<button onclick="intentarEnviarCola()">Reintentar todos ahora</button>' : '');
    }

    let enviandoCola = false;   // evita que dos disparadores a la vez (evento "online" + el intervalo
                                 // de siempre, algo comun cuando la Wi-Fi entra y sale) reenvien el mismo
                                 // pedido guardado dos veces antes de que el primero termine de limpiarlo
    async function intentarEnviarCola() {
      if (enviandoCola) return;
      let cola = leerCola();
      if (cola.length === 0) return;
      enviandoCola = true;
      const restante = [];
      for (const entrada of cola) {
        try {
          const res = await fetch('/api/pedidos', { method: 'POST', body: JSON.stringify(entrada.payload) });
          const data = await res.json().catch(() => null);
          if (!res.ok) {
            entrada.error = (data && data.error) ? data.error : 'El servidor rechazo el pedido.';
            restante.push(entrada);
          }
          // si res.ok, se envio bien y se descarta de la cola
        } catch (e) {
          // sigue sin conexion, se queda para el proximo intento
          restante.push(entrada);
        }
      }
      guardarCola(restante);
      enviandoCola = false;
      cargarCatalogo();
      cargarMisPedidos();
    }

    window.addEventListener('online', intentarEnviarCola);
    setInterval(intentarEnviarCola, 10000);
    actualizarColaUI();

    // ---- Mis pedidos de hoy ----
    function hoyStr() {
      const d = new Date();
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }

    // Firma de la lista mostrada: en el refresco automatico solo se redibuja
    // si algo cambio de verdad (un pedido nuevo, o la PC cobro/cancelo uno).
    // Asi no se cierra el selector de metodo de pago mientras se usa, pero
    // en cuanto la caja cobra un pedido, su menu desaparece solo.
    let ultimaFirmaMP = null;
    const expandidosMP = new Set();   // pedidos con la lista de productos desplegada

    function abrirMisPedidos(abrir) {
      document.getElementById('misPedidos').style.display = abrir ? 'block' : 'none';
      document.getElementById('flechaMisPedidos').innerHTML = abrir ? '&#9652;' : '&#9662;';
    }
    function toggleMisPedidos() {
      abrirMisPedidos(document.getElementById('misPedidos').style.display === 'none');
    }
    function toggleItemsMP(id) {
      if (expandidosMP.has(id)) expandidosMP.delete(id); else expandidosMP.add(id);
      const abierto = expandidosMP.has(id);
      const caja = document.getElementById('mp-items-' + id);
      const btn = document.getElementById('mp-btnitems-' + id);
      if (caja) caja.style.display = abierto ? 'block' : 'none';
      if (btn) btn.innerHTML = abierto ? 'Ocultar productos &#9652;' : 'Ver productos &#9662;';
    }

    function firmaMP(p) {
      return [p.id, p.estado, p.metodoPago || '', p.totalCobrado, p.totalProductos].join('|');
    }

    async function cargarMisPedidos(auto) {
      if (posActivo()) { return cargarVentasPOS(auto); }
      const cont = document.getElementById('misPedidos');
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) {
        ultimaFirmaMP = null;
        cont.innerHTML = '<div style="color:#94a3b8; font-size:13px;">Escribe tu nombre para ver tus pedidos.</div>';
        return;
      }
      const claveCache = 'misPedidosCache_' + nombre.toLowerCase();
      try {
        const res = await fetch('/api/pedidos');
        const pedidos = await res.json();
        const hoy = hoyStr();
        const mios = pedidos.filter(p =>
          (p.vendedor || '').trim().toLowerCase() === nombre.toLowerCase() && (p.hora || '').startsWith(hoy)
        ).sort((a, b) => b.id - a.id);
        // Se guarda localmente lo ultimo visto de este vendedor hoy, para
        // poder mostrarlo si mas tarde se pierde la conexion con la PC
        // (por ejemplo, saliendo a buscar la mercancia del pedido).
        try { localStorage.setItem(claveCache, JSON.stringify({ dia: hoy, mios })); } catch(e) {}
        const pend = mios.filter(p => p.estado === 'pendiente').length;
        document.getElementById('resumenMisPedidos').textContent = mios.length === 0 ? '' :
          ' · ' + mios.length + (mios.length === 1 ? ' pedido' : ' pedidos') +
          (pend ? ' · ' + pend + (pend === 1 ? ' pendiente' : ' pendientes') : '');
        const firma = nombre.toLowerCase() + '#' + mios.map(firmaMP).join(',');
        if (auto === true && firma === ultimaFirmaMP) return;
        ultimaFirmaMP = firma;
        cont.innerHTML = mios.length === 0
          ? '<div style="color:#94a3b8; font-size:13px;">Sin pedidos todavia hoy.</div>'
          : mios.map(renderMiPedido).join('');
      } catch (e) {
        // Sin conexion con la PC: se muestra lo ultimo que se sabia de este
        // vendedor hoy (guardado la ultima vez que si hubo conexion, por
        // ejemplo justo despues de enviar este mismo pedido), en vez de
        // dejar la pantalla en blanco hasta que vuelva a conectar.
        ultimaFirmaMP = null;
        let cache = null;
        try { cache = JSON.parse(localStorage.getItem(claveCache) || 'null'); } catch(e2) {}
        if (cache && cache.dia === hoyStr() && Array.isArray(cache.mios)) {
          cont.innerHTML =
            '<div style="color:#92400e; background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:8px; font-size:12px; margin-bottom:8px;">⚠ Sin conexion con la PC ahora mismo — esto es lo ultimo que se sabe (puede que la caja ya haya cobrado o cambiado algo mientras tanto).</div>' +
            (cache.mios.length === 0
              ? '<div style="color:#94a3b8; font-size:13px;">Sin pedidos todavia hoy.</div>'
              : cache.mios.map(renderMiPedido).join(''));
        } else {
          cont.innerHTML = '<div style="color:#991b1b; font-size:13px;">No se pudo cargar tus pedidos (revisa la conexion).</div>';
        }
      }
    }

    // Metodo elegido por pedido pendiente: se recuerda para que no se
    // reinicie a "Efectivo" cada vez que se refresca la lista.
    const metodoMP = {};

    function totalBasePedido(p) {
      return Number(p.totalProductos !== undefined && p.totalProductos !== null ? p.totalProductos : (p.total || 0));
    }

    function textoTotalMP(base, metodo) {
      return metodo === 'Transferencia'
        ? ('A cobrar por transferencia (x2): $' + (base * 2).toFixed(2))
        : ('A cobrar: $' + base.toFixed(2));
    }

    // Mientras la caja no cobre, el vendedor puede cambiar el metodo de pago:
    // se guarda en el servidor para que la PC vea el metodo y el monto nuevos.
    async function cambiarMetodoMP(id, base) {
      const sel = document.getElementById('mp-metodo-' + id);
      if (!sel) return;
      const nuevo = sel.value;
      metodoMP[id] = nuevo;
      const el = document.getElementById('mp-total-' + id);
      if (el) el.textContent = textoTotalMP(base, nuevo);
      try {
        const res = await fetch('/api/pedidos/' + id + '/metodo', { method: 'POST', body: JSON.stringify({ metodoPago: nuevo, pin: miPin() }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          alert((data && data.error) || 'No se pudo cambiar el metodo de pago.');
          delete metodoMP[id];
          cargarMisPedidos();
        }
      } catch (e) {
        alert('Sin conexion con la PC: el cambio de metodo NO se guardo. Vuelve a intentarlo.');
        delete metodoMP[id];
        cargarMisPedidos();
      }
    }

    function renderMiPedido(p) {
      const totalBase = totalBasePedido(p);
      const total = Number(p.totalCobrado !== undefined && p.totalCobrado !== null ? p.totalCobrado : totalBase);
      const horaCorta = (p.hora || '').slice(11);
      let totalTxt = '$' + total.toFixed(2);
      if (p.estado === 'cobrado' && p.metodoPago) totalTxt += ' — ' + p.metodoPago;
      let acciones = '<button class="mp-btn-imprimir" onclick="imprimirMP(' + p.id + ')">Reimprimir</button>';
      if (p.estado === 'pendiente') {
        const elegido = metodoMP[p.id] || p.metodoPago || 'Efectivo';
        const opt = (v, t) => '<option value="' + v + '"' + (elegido === v ? ' selected' : '') + '>' + t + '</option>';
        totalTxt = textoTotalMP(totalBase, elegido);
        acciones =
          '<select class="mp-select" id="mp-metodo-' + p.id + '" onchange="cambiarMetodoMP(' + p.id + ', ' + totalBase + ')">' +
            opt('Efectivo', 'Efectivo') +
            opt('Transferencia', 'Transferencia (x2)') +
            opt('Otro', 'Otro') +
          '</select>' +
          '<button class="mp-btn-cobrar" onclick="cobrarMP(' + p.id + ', ' + totalBase + ')">Cobrar</button>' +
          '<button class="mp-btn-editar" onclick="editarMP(' + p.id + ')">Editar</button>' +
          '<button class="mp-btn-cancelar" onclick="cancelarMP(' + p.id + ')">Cancelar</button>' +
          acciones;
      }
      // Lista de lo que se anoto en el pedido (para rectificar si se quedo algo)
      const abierto = expandidosMP.has(p.id);
      const lineas = (p.items || []).map(it =>
        '<div><span>' + it.cantidad + ' x ' + escaparHtml(it.nombre) + '</span><span>$' + (it.precio * it.cantidad).toFixed(2) + '</span></div>'
      ).join('');
      const itemsHtml =
        '<button class="mp-btn-items" id="mp-btnitems-' + p.id + '" onclick="toggleItemsMP(' + p.id + ')">' +
          (abierto ? 'Ocultar productos &#9652;' : 'Ver productos &#9662;') + '</button>' +
        '<div class="mp-items" id="mp-items-' + p.id + '" style="display:' + (abierto ? 'block' : 'none') + ';">' + lineas +
          '<div style="border-top:1px solid #e2e8f0; margin-top:4px; padding-top:5px; font-weight:700;"><span>Total productos</span><span>$' + totalBase.toFixed(2) + '</span></div>' +
        '</div>';
      return '<div class="mp-card" data-sig="' + firmaMP(p) + '">' +
        '<div class="mp-top"><span>Folio #' + p.id + ' — ' + horaCorta + '</span><span class="mp-estado ' + p.estado + '">' + p.estado.toUpperCase() + '</span></div>' +
        '<div class="mp-total" id="mp-total-' + p.id + '">' + totalTxt + '</div>' +
        (p.nota ? '<div style="font-size:12px; color:#0369a1; margin-bottom:6px;">Tu nota: ' + escaparHtml(p.nota) + '</div>' : '') +
        itemsHtml +
        '<div class="mp-acciones">' + acciones + '</div>' +
      '</div>';
    }

    async function cobrarMP(id, totalBase) {
      try {
        const sel = document.getElementById('mp-metodo-' + id);
        const metodoPago = sel ? sel.value : 'Efectivo';
        const totalCobrar = metodoPago === 'Transferencia' ? totalBase * 2 : totalBase;
        const totalTxt = (typeof totalCobrar === 'number' && !isNaN(totalCobrar)) ? totalCobrar.toFixed(2) : '';
        const entrada = prompt('Monto recibido del cliente' + (totalTxt ? ' (total $' + totalTxt + ')' : '') + ':', totalTxt);
        if (entrada === null) return; // cancelo
        const montoRecibido = parseFloat(String(entrada).replace(',', '.'));
        if (isNaN(montoRecibido) || montoRecibido < 0) { alert('Escribe un monto valido.'); return; }
        if (montoRecibido < totalCobrar) {
          alert('El monto recibido ($' + montoRecibido.toFixed(2) + ') es menor que el total a cobrar ($' + totalCobrar.toFixed(2) + '). Revisa el monto.');
          return;
        }
        const res = await fetch('/api/pedidos/' + id + '/cobrar', { method: 'POST', body: JSON.stringify({ metodoPago, montoRecibido, pin: miPin() }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          // Por ejemplo: la caja ya lo cobro, o el PIN no coincide. Se avisa y se actualiza la lista.
          alert((data && data.error) || 'No se pudo cobrar el pedido.');
          if (data && data.requierePin && !autoservicioActivo) abrirLogin((nombreInput.value || '').trim());
        } else {
          delete metodoMP[id];
        }
        cargarMisPedidos();
      } catch (e) { alert('No se pudo cobrar (revisa la conexion con la PC).'); }
    }

    async function imprimirMP(id) {
      try {
        const res = await fetch('/api/pedidos/' + id + '/imprimir', { method: 'POST' });
        const data = await res.json().catch(() => null);
        if (!data || !data.ok) alert('No se pudo imprimir: ' + ((data && data.error) || 'revisa la impresora en la PC'));
      } catch (e) { alert('No se pudo imprimir (revisa la conexion con la PC).'); }
    }

    async function cancelarMP(id) {
      if (!confirm('¿Cancelar este pedido? El stock de sus productos se devuelve.')) return;
      try {
        const res = await fetch('/api/pedidos/' + id + '/cancelar', { method: 'POST', body: JSON.stringify({ pin: miPin() }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          alert((data && data.error) || 'No se pudo cancelar el pedido.');
          if (data && data.requierePin && !autoservicioActivo) abrirLogin((nombreInput.value || '').trim());
        }
        cargarMisPedidos();
        cargarCatalogo();
      } catch (e) { alert('No se pudo cancelar (revisa la conexion con la PC).'); }
    }

    async function editarMP(id) {
      if (!confirm('Esto cancela el pedido #' + id + ' (devuelve su stock) y carga sus productos aqui para que los ajustes y lo reenvies. ¿Continuar?')) return;
      try {
        const res = await fetch('/api/pedidos');
        const pedidos = await res.json();
        const pedido = pedidos.find(p => p.id === id);
        if (!pedido) { alert('No se encontro el pedido (puede que ya lo hayan cambiado en la PC).'); return; }
        if (pedido.estado !== 'pendiente') {
          alert('Este pedido ya fue ' + (pedido.estado === 'cobrado' ? 'cobrado' : 'cancelado') + ' y no se puede editar.');
          cargarMisPedidos();
          return;
        }
        const resC = await fetch('/api/pedidos/' + id + '/cancelar', { method: 'POST', body: JSON.stringify({ pin: miPin() }) });
        const datC = await resC.json().catch(() => null);
        if (!resC.ok || !datC || !datC.ok) {
          // Si no se pudo anular (p. ej. la caja ya lo cobro, o el PIN no coincide) NO se
          // carga al carrito: se duplicaria la venta.
          alert((datC && datC.error) || 'No se pudo anular el pedido para editarlo.');
          if (datC && datC.requierePin && !autoservicioActivo) abrirLogin((nombreInput.value || '').trim());
          cargarMisPedidos();
          return;
        }
        await cargarCatalogo();
        carrito = (pedido.items || []).map(it => {
          const enCatalogo = catalogo.find(p => p.sku === it.sku);
          return { sku: it.sku, nombre: it.nombre, precio: it.precio, cantidad: it.cantidad, stock: enCatalogo ? enCatalogo.stock : null };
        });
        const selMetodo = document.getElementById('metodoPago');
        if (selMetodo && pedido.metodoPago) selMetodo.value = pedido.metodoPago;
        renderCarrito();
        cargarMisPedidos();
        document.getElementById('accionesPostEnvio').style.display = 'none';
        mostrarMensaje('Pedido cargado en el carrito. Ajustalo y toca "Enviar pedido".', true);
        const destino = document.getElementById('seccionPedidoActual');
        if (destino) destino.scrollIntoView({ behavior: 'smooth' });
      } catch (e) { alert('No se pudo editar (revisa la conexion con la PC).'); }
    }

    nombreInput.addEventListener('change', cargarMisPedidos);
    cargarMisPedidos();
    setInterval(() => cargarMisPedidos(true), 8000);
    nombreInput.addEventListener('change', revisarPedidosAsignados);

    // ---- Permisos que la PC le da a este vendedor (que puede hacer y a donde entrar) ----
    window.__permisos = {};
    let permisosFirma = '';
    const CLAVES_PERMISOS = ['crearPedidos', 'cobrar', 'cancelar', 'editar', 'imprimir', 'descuentos', 'verStock', 'misPedidos', 'ajustes', 'modoCliente', 'asignados', 'notificaciones', 'mensajes'];
    function puede(k) { return !window.__permisos || window.__permisos[k] !== false; }
    function aplicarPermisos(p) {
      window.__permisos = p || {};
      CLAVES_PERMISOS.forEach(k => document.body.classList.toggle('sin-' + k, window.__permisos[k] === false));
      if (!puede('cobrar')) {
        const r = document.querySelector('input[name="estadoPago"][value="pendiente"]');
        if (r && !r.checked) { r.checked = true; try { toggleCobro(); } catch (e) {} }
      }
      try { renderCarrito(); } catch (e) {}
      try { buscar(); } catch (e) {}
    }
    async function revisarPermisos() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) {
        if (permisosFirma !== '') { permisosFirma = ''; aplicarPermisos({}); }
        return;
      }
      try {
        const res = await fetch('/api/permisos?vendedor=' + encodeURIComponent(nombre));
        const d = await res.json();
        if (!d || !d.permisos) return;
        const firma = nombre + '|' + JSON.stringify(d.permisos);
        if (firma !== permisosFirma) { permisosFirma = firma; aplicarPermisos(d.permisos); }
      } catch (e) {}
    }
    revisarPermisos();
    setInterval(revisarPermisos, 8000);
    nombreInput.addEventListener('change', revisarPermisos);

    // ---- Presencia (para que la PC sepa que este vendedor esta conectado) ----
    async function avisarConectado() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre || autoservicioActivo) return;   // un cliente NO cuenta como vendedor en linea
      try { await fetch('/api/vendedores/ping', { method: 'POST', body: JSON.stringify({ nombre }) }); } catch (e) {}
    }
    avisarConectado();
    setInterval(avisarConectado, 7000);

    // ---- Pedidos que la PC arma y manda a este vendedor ----
    function beepAsignado() {
      try {
        const ctx = new (window.AudioContext || window.webkitAudioContext)();
        const tocar = (inicio) => {
          const o = ctx.createOscillator();
          const g = ctx.createGain();
          o.connect(g); g.connect(ctx.destination);
          o.type = 'square';
          o.frequency.setValueAtTime(880, ctx.currentTime + inicio);
          g.gain.setValueAtTime(0.0001, ctx.currentTime + inicio);
          g.gain.exponentialRampToValueAtTime(0.9, ctx.currentTime + inicio + 0.02);
          o.start(ctx.currentTime + inicio);
          o.frequency.setValueAtTime(1200, ctx.currentTime + inicio + 0.15);
          g.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + inicio + 0.5);
          o.stop(ctx.currentTime + inicio + 0.5);
        };
        tocar(0);
        tocar(0.55);
      } catch (e) {}
    }

    // ---- Campanita de notificaciones (todo lo que la caja le manda al
    // vendedor: pedidos armados y alertas de descuadre) ----
    let notificacionesCampana = [];
    let notiCampanaNoLeidas = 0;

    function horaCorta12() {
      try { return new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }); }
      catch (e) { return ''; }
    }

    function agregarNotificacionCampana(texto) {
      notificacionesCampana.unshift({ texto, hora: horaCorta12() });
      if (notificacionesCampana.length > 30) notificacionesCampana.length = 30;
      notiCampanaNoLeidas++;
      actualizarBadgeCampana();
      renderCampana();
    }

    function actualizarBadgeCampana() {
      const b = document.getElementById('badgeCampana');
      if (!b) return;
      if (notiCampanaNoLeidas > 0) {
        b.style.display = 'flex';
        b.textContent = notiCampanaNoLeidas > 9 ? '9+' : String(notiCampanaNoLeidas);
      } else {
        b.style.display = 'none';
      }
    }

    function renderCampana() {
      const cont = document.getElementById('listaCampana');
      if (!cont) return;
      cont.innerHTML = notificacionesCampana.length === 0
        ? '<div class="campana-vacio">Sin notificaciones.</div>'
        : notificacionesCampana.map(n =>
            '<div class="campana-item"><div class="campana-hora">' + n.hora + '</div><div>' + escaparHtml(n.texto) + '</div></div>'
          ).join('');
    }

    function toggleCampana() {
      const p = document.getElementById('panelCampana');
      if (!p) return;
      const abrir = p.style.display !== 'block';
      p.style.display = abrir ? 'block' : 'none';
      if (abrir) { notiCampanaNoLeidas = 0; actualizarBadgeCampana(); }
    }
    renderCampana();

    // Pedidos que la caja arma y manda a este vendedor: NO se suman de una vez
    // al carrito que el vendedor ya este armando. Se guardan aparte, en
    // "Pedidos que te armo la caja", y el vendedor decide cuando agregarlos
    // a su pedido actual (por ejemplo, cuando termine el que tiene entre
    // manos). Al abrirlos se avisa a la caja que ya se vieron.
    let asignadosPendientes = [];

    function renderAsignados() {
      const cont = document.getElementById('asignadosLista');
      const sec = document.getElementById('seccionAsignados');
      if (asignadosPendientes.length === 0) {
        sec.style.display = 'none';
        cont.innerHTML = '';
        return;
      }
      sec.style.display = 'block';
      cont.innerHTML = asignadosPendientes.map(a => {
        const total = (a.items || []).reduce((s, it) => s + it.precio * it.cantidad, 0);
        const horaCorta = (a.hora || '').slice(11);
        const lineas = (a.items || []).map(it =>
          '<div><span>' + it.cantidad + ' x ' + escaparHtml(it.nombre) + '</span><span>$' + (it.precio * it.cantidad).toFixed(2) + '</span></div>'
        ).join('');
        const esTodos = !!a.todos;
        const btnVisto = a.visto
          ? '<button class="mp-btn-visto" disabled>' + (esTodos ? 'Lo tomaste &#10003;' : 'Visto &#10003;') + '</button>'
          : (esTodos
              ? '<button class="mp-btn-visto" onclick="tomarAsignado(' + a.id + ')">Yo lo tomo</button>'
              : '<button class="mp-btn-visto" onclick="marcarVistoAsignado(' + a.id + ')">Ya lo vi</button>');
        // Si es "para todos" y todavia no lo tomaste, no dejamos agregarlo al
        // carrito hasta que confirmes que lo tomas (para que no lo agreguen
        // dos vendedores a la vez).
        const btnAgregar = (!esTodos || a.visto)
          ? '<button class="mp-btn-agregar-asig" onclick="agregarAsignadoAlCarrito(' + a.id + ')">Agregar a mi pedido</button>'
          : '';
        const etiquetaOrigen = a.cliente
          ? ('Cliente: ' + escaparHtml(a.cliente) + ' — ')
          : (esTodos ? 'Pedido para el primero que lo tome — ' : 'De la caja — ');
        return '<div class="asig-card">' +
          '<div class="mp-top"><span>' + etiquetaOrigen + horaCorta + '</span></div>' +
          '<div class="mp-items" style="display:block;">' + lineas + '</div>' +
          (a.nota ? '<div style="font-size:13px; color:#0369a1; margin:4px 0;">Nota de la caja: ' + escaparHtml(a.nota) + '</div>' : '') +
          '<div class="mp-total">Total: $' + total.toFixed(2) + '</div>' +
          '<div class="mp-acciones">' + btnVisto + btnAgregar +
          '</div>' +
        '</div>';
      }).join('');
    }

    async function marcarVistoAsignado(id) {
      const a = asignadosPendientes.find(x => x.id === id);
      if (a) a.visto = true;
      renderAsignados();
      try { await fetch('/api/pedidos/asignados/' + id + '/visto', { method: 'POST' }); } catch (e) {}
    }

    async function tomarAsignado(id) {
      const nombre = (nombreInput.value || '').trim();
      try {
        const res = await fetch('/api/pedidos/asignados/' + id + '/tomar', { method: 'POST', body: JSON.stringify({ vendedor: nombre }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          const quien = (data && data.tomadoPor) ? data.tomadoPor : 'otro vendedor';
          mostrarMensaje('Ese pedido ya lo tomo ' + quien + '.', false);
          const idx = asignadosPendientes.findIndex(x => x.id === id);
          if (idx !== -1) asignadosPendientes.splice(idx, 1);
          renderAsignados();
          return;
        }
        const a = asignadosPendientes.find(x => x.id === id);
        if (a) { a.visto = true; a.tomadoPor = nombre; a.horaTomado = new Date().toISOString(); }
        renderAsignados();
        mostrarMensaje('Lo tomaste. Toca "Agregar a mi pedido" cuando lo vayas a armar.', true);
      } catch (e) { mostrarMensaje('No se pudo confirmar (revisa la conexion).', false); }
    }

    function agregarAsignadoAlCarrito(id) {
      const idx = asignadosPendientes.findIndex(x => x.id === id);
      if (idx === -1) return;
      const a = asignadosPendientes[idx];
      if (a.todos && !a.visto) { tomarAsignado(id); return; }
      for (const it of (a.items || [])) {
        const existente = carrito.find(i => i.sku === it.sku);
        if (existente) existente.cantidad += Number(it.cantidad) || 0;
        else carrito.push({ sku: it.sku, nombre: it.nombre, precio: it.precio, cantidad: Number(it.cantidad) || 0, stock: null });
      }
      if (a.todos) {
        origenesAsignadosCarrito.push({ asignadoId: a.id, tomadoPor: a.tomadoPor || (nombreInput.value || '').trim(), horaTomado: a.horaTomado || new Date().toISOString(), cliente: a.cliente || '' });
      }
      if (!a.visto) marcarVistoAsignado(id);
      asignadosPendientes.splice(idx, 1);
      renderCarrito();
      renderAsignados();
      mostrarMensaje('Agregado a tu pedido actual.', true);
      const destino = document.getElementById('seccionPedidoActual');
      if (destino) destino.scrollIntoView({ behavior: 'smooth' });
    }

    function irAAsignados() {
      document.getElementById('bannerAsignado').style.display = 'none';
      const destino = document.getElementById('seccionAsignados');
      if (destino) destino.scrollIntoView({ behavior: 'smooth' });
    }

    async function revisarPedidosAsignados() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) return;
      try {
        const res = await fetch('/api/pedidos/asignados?vendedor=' + encodeURIComponent(nombre));
        const data = await res.json();
        const nuevos = (data && data.pedidos) || [];
        if (nuevos.length === 0) return;
        let hayNuevo = false;
        for (const a of nuevos) {
          if (asignadosPendientes.some(x => x.id === a.id)) continue;   // ya lo tiene: no se repite la tarjeta ni el aviso
          hayNuevo = true;
          asignadosPendientes.push({ id: a.id, items: a.items || [], hora: a.hora, todos: !!a.todos, visto: false, cliente: a.cliente || '', nota: a.nota || '' });
          const nProd = (a.items || []).length;
          const quePedido = a.cliente
            ? ('El cliente "' + a.cliente + '" mando un pedido (')
            : (a.todos ? 'La caja mando un pedido para todos (' : 'La caja te armo un pedido nuevo (');
          agregarNotificacionCampana(quePedido + nProd + (nProd === 1 ? ' producto' : ' productos') + ').');
        }
        if (!hayNuevo) return;
        renderAsignados();
        beepAsignado();
        document.getElementById('bannerAsignado').style.display = 'block';
        mostrarMensaje('Hay un pedido nuevo: revisalo abajo, en "Pedidos que te armo la caja".', true);
      } catch (e) {}
    }

    // Los pedidos "para todos" que ya recibiste pero no has tomado pueden
    // desaparecer porque otro vendedor los tomo primero. Tambien revisa
    // cualquier pedido asignado (para todos o directo) que la caja haya
    // retirado por error, antes de que lo tomaras. Sin esto, la tarjeta se
    // quedaba en pantalla hasta que recargabas la pagina.
    async function revisarAsignadosTomadosPorOtro() {
      if (asignadosPendientes.length === 0) return;
      const nombre = (nombreInput.value || '').trim();
      try {
        const res = await fetch('/api/pedidos/asignados/estado');
        const data = await res.json();
        const estados = (data && data.asignados) || [];
        let cambio = false;
        for (const a of [...asignadosPendientes]) {
          const e = estados.find(x => x.id === a.id);
          if (!e) continue;
          const loTomoOtro = a.todos && !a.visto && e.tomadoPor && e.tomadoPor !== nombre;
          if (loTomoOtro || e.retirado) {
            const idx = asignadosPendientes.findIndex(x => x.id === a.id);
            if (idx !== -1) { asignadosPendientes.splice(idx, 1); cambio = true; }
            if (e.retirado) mostrarMensaje('La caja retiro un pedido que te habia armado.', false);
          }
        }
        if (cambio) renderAsignados();
      } catch (e) {}
    }
    setInterval(revisarAsignadosTomadosPorOtro, 5000);

    // ---- Avisos push locales: la caja avisa de precios, stock, anulaciones y permisos ----
    async function mostrarNotificacionSistema(titulo, cuerpo) {
      try {
        if (!('Notification' in window) || Notification.permission !== 'granted') return;
        const opts = { body: cuerpo, icon: '/icon-192.png', badge: '/icon-192.png', tag: 'tt-' + Date.now(), vibrate: [200, 100, 200] };
        let reg = null;
        try { reg = await navigator.serviceWorker.getRegistration(); } catch (e) {}
        if (reg && reg.showNotification) await reg.showNotification(titulo, opts);
        else new Notification(titulo, opts);
      } catch (e) {}
    }
    async function activarAvisosTelefono() {
      const st = document.getElementById('estadoAvisosTel');
      if (!('Notification' in window)) {
        st.textContent = 'Este navegador no permite avisos del sistema. Igual te avisaremos con sonido y vibración dentro de la app.';
        return;
      }
      if (!window.isSecureContext) {
        st.textContent = 'Chrome bloquea los avisos del sistema en direcciones http. Para activarlos abre chrome://flags/#unsafely-treat-insecure-origin-as-secure, escribe ' + location.origin + ', activa la opción y reinicia Chrome. Mientras tanto, con la app abierta te avisamos con sonido, vibración y la campanita.';
        return;
      }
      const r = await Notification.requestPermission();
      st.textContent = r === 'granted' ? 'Avisos del teléfono activados.' : 'Los avisos están bloqueados: actívalos en los ajustes del navegador para este sitio.';
      if (r === 'granted') mostrarNotificacionSistema('Toto Tools', 'Avisos activados en este teléfono.');
    }
    // ================= DESCARGAR LO QUE VA QUEDANDO (stock actual) =================
    // Baja un Excel con el catalogo y la cantidad que queda de cada producto (ya descontadas
    // las ventas). Usa los mismos encabezados que el Excel de la PC, asi se puede volver a
    // cargar como catalogo. Funciona tambien sin conexion (usa el catalogo guardado en el telefono).
    async function descargarStockActual() {
      const estado = document.getElementById('estadoDescargaExcel');
      try {
        if (!catalogo || !catalogo.length) { estado.textContent = 'Todavia no hay catalogo cargado en este telefono.'; return; }
        if (typeof XLSX === 'undefined') { estado.textContent = 'No se pudo preparar el Excel (falta la libreria). Recarga la pagina.'; return; }
        estado.textContent = 'Preparando archivo...';
        let m = {};
        try { const rc = await fetch('/api/catalogo/config'); const c = await rc.json(); m = (c && c.mapeo) || {}; } catch (e) {}
        const hSku = m.sku || 'SKU', hNom = m.nombre || 'Nombre', hPre = m.precio || 'Precio', hStk = m.stock || 'Cantidad';
        const filas = catalogo.map(p => {
          const f = {};
          f[hSku] = p.sku || '';
          f[hNom] = p.nombre;
          f[hPre] = p.precio;
          f[hStk] = (p.stock === null || p.stock === undefined) ? '' : p.stock;
          return f;
        });
        const ws = XLSX.utils.json_to_sheet(filas, { header: [hSku, hNom, hPre, hStk] });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Stock actual');
        const datos = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([datos], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const d = new Date();
        const nombre = 'stock_actual_' + hoyStr() + '_' + String(d.getHours()).padStart(2, '0') + String(d.getMinutes()).padStart(2, '0') + '.xlsx';
        const como = await guardarArchivoEnMovil(blob, nombre);
        estado.textContent = como === 'compartido'
          ? 'Listo: elige donde guardar o con que abrir "' + nombre + '".'
          : 'Descargado: ' + nombre;
      } catch (e) {
        if (e && e.message === 'sin-soporte') {
          estado.textContent = 'Este telefono no deja guardar el archivo desde la app. Abre ' + location.origin + '/vendedor en Chrome y descargalo ahi (o actualiza la app).';
        } else if (e && (e.name === 'AbortError' || /cancel/i.test(String(e.message || '')))) {
          estado.textContent = 'Cancelado.';
        } else {
          estado.textContent = 'No se pudo preparar el archivo (' + ((e && e.message) || 'error') + ').';
        }
      }
    }

    // ================= MODO PUNTO DE VENTA (lo activa el administrador con su clave) =================
    // Con el modo activo, este telefono: cobra por defecto al enviar, permite descuentos por producto,
    // y "Ventas (caja)" muestra el historial, el reporte de efectivo / transferencia y las devoluciones.
    // Todo sigue pasando por la PC, asi que las dos pantallas ven lo mismo.
    var posUltimaFirma = null;
    var posPedidosPorClave = {};

    function posToken() { return localStorage.getItem('posToken') || ''; }
    function posActivo() {
      const n = (nombreInput.value || '').trim().toLowerCase();
      return !autoservicioActivo && !!posToken() && !!n && localStorage.getItem('posVendedor') === n;
    }
    function posLimpiarLocal() { localStorage.removeItem('posToken'); localStorage.removeItem('posVendedor'); }
    function descuentosDisponibles() { return permitirDescuentosActual || posActivo(); }
    function dineroPOS(n) { return '$' + (Number(n) || 0).toFixed(2); }

    function aplicarEstadoPagoPorDefecto() {
      const quiere = (posActivo() && puede('cobrar')) ? 'cobrado' : 'pendiente';
      const r = document.querySelector('input[name=estadoPago][value=' + quiere + ']');
      if (r) { r.checked = true; toggleCobro(); }
    }

    function posAplicarUI(reiniciarEstado) {
      const on = posActivo();
      document.body.classList.toggle('modoPOS', on);
      document.getElementById('posInactivoBox').style.display = on ? 'none' : 'block';
      document.getElementById('posActivoBox').style.display = on ? 'block' : 'none';
      document.getElementById('tituloMisPedidos').textContent = on ? 'Ventas (caja)' : 'Mis pedidos de hoy';
      posUltimaFirma = null;
      abrirMisPedidos(document.getElementById('misPedidos').style.display !== 'none');
      if (reiniciarEstado) aplicarEstadoPagoPorDefecto();
      try { renderCarrito(); } catch (e) {}
    }

    async function activarModoPOS() {
      const msg = document.getElementById('posMensaje');
      const clave = (document.getElementById('posClaveInput').value || '').trim();
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) { msg.textContent = 'Entra primero como vendedor.'; return; }
      if (!clave) { msg.textContent = 'Escribe la clave de administrador.'; return; }
      msg.textContent = 'Comprobando...';
      try {
        const res = await fetch('/api/pos/activar', { method: 'POST', body: JSON.stringify({ vendedor: nombre, pin: miPin(), clave: clave }) });
        const d = await res.json().catch(() => null);
        if (!res.ok || !d || !d.ok) {
          msg.textContent = (d && d.error) || 'No se pudo activar.';
          if (d && d.requierePin) abrirLogin(nombre);
          return;
        }
        localStorage.setItem('posToken', d.token);
        localStorage.setItem('posVendedor', nombre.toLowerCase());
        document.getElementById('posClaveInput').value = '';
        msg.textContent = '';
        cerrarMenu();
        posAplicarUI(true);
        abrirMisPedidos(true);
      } catch (e) {
        msg.textContent = 'Sin conexion con la PC: hace falta conexion para activarlo.';
      }
    }

    function desactivarModoPOS() {
      posLimpiarLocal();
      posAplicarUI(true);
      document.getElementById('posMensaje').textContent = 'Modo punto de venta desactivado.';
      cargarMisPedidos();
    }

    // Si el administrador cambia o quita la clave en la PC, este telefono pierde el modo.
    async function posVerificar() {
      const on = posActivo();
      if (on !== document.body.classList.contains('modoPOS')) posAplicarUI(true);
      if (!on) return;
      try {
        const r = await fetch('/api/pos/estado');
        const d = await r.json().catch(() => null);
        if (d && d.ok && !d.activo) {
          posLimpiarLocal();
          posAplicarUI(true);
          cargarMisPedidos();
          mostrarMensaje('El administrador desactivo el modo punto de venta en este telefono.', false);
        }
      } catch (e) {}   // sin conexion: se conserva el modo
    }

    async function cargarVentasPOS(auto) {
      const cont = document.getElementById('misPedidos');
      if (auto === true && cont.style.display === 'none') return;   // cerrado: no se gasta red ni se carga la PC
      const dias = document.getElementById('posRango').value || '1';
      const quien = document.getElementById('posQuien').value || 'todos';
      const claveCache = 'posVentasCache_' + (nombreInput.value || '').trim().toLowerCase() + '_' + dias + '_' + quien;
      try {
        const res = await fetch('/api/pos/ventas?dias=' + encodeURIComponent(dias) + '&quien=' + encodeURIComponent(quien));
        const d = await res.json().catch(() => null);
        if (res.status === 403 && d && d.posInactivo) {
          posLimpiarLocal();
          posAplicarUI(true);
          mostrarMensaje('El modo punto de venta ya no esta activo en este telefono.', false);
          return cargarMisPedidos();
        }
        if (!res.ok || !d || !d.ok) throw new Error((d && d.error) || 'error');
        // Pedidos propios que siguen pendientes de cobro (la caja o el vendedor los cobra con los botones de siempre).
        let pend = [];
        try {
          const rp = await fetch('/api/pedidos');
          const todos = await rp.json();
          const yo = (nombreInput.value || '').trim().toLowerCase();
          pend = (Array.isArray(todos) ? todos : []).filter(p => p.estado === 'pendiente' && (p.vendedor || '').trim().toLowerCase() === yo && (p.hora || '').startsWith(hoyStr())).sort((a, b) => b.id - a.id);
        } catch (e3) {}
        try { localStorage.setItem(claveCache, JSON.stringify({ t: Date.now(), d: d, pend: pend })); } catch (e) {}
        pintarVentasPOS(d, 0, auto, pend);
      } catch (e) {
        // Sin conexion con la PC: se muestra lo ultimo que se vio de este mismo periodo.
        posUltimaFirma = null;
        let cache = null;
        try { cache = JSON.parse(localStorage.getItem(claveCache) || 'null'); } catch (e2) {}
        if (cache && cache.d) pintarVentasPOS(cache.d, cache.t, false, cache.pend || []);
        else cont.innerHTML = '<div style="color:#991b1b; font-size:13px;">No se pudo cargar el historial (revisa la conexion con la PC).</div>';
      }
    }

    function pintarVentasPOS(d, cacheDe, auto, pend) {
      pend = pend || [];
      const cont = document.getElementById('misPedidos');
      const r = d.reporte || {};
      const pedidos = d.pedidos || [];
      document.getElementById('resumenMisPedidos').textContent = ' | ' + (r.ventas || 0) + (r.ventas === 1 ? ' venta' : ' ventas') + ' | neto ' + dineroPOS(r.netoTotal);
      const firma = JSON.stringify([r, pedidos.map(p => [p.id, p.hora, p.totalCobrado, p.devuelto]), cacheDe || 0, pend.map(firmaMP)]);
      if (auto === true && firma === posUltimaFirma) return;
      posUltimaFirma = firma;
      posPedidosPorClave = {};
      const fila = (t, v, fuerte) => '<div' + (fuerte ? ' style="font-weight:700;"' : '') + '><span>' + t + '</span><span>' + v + '</span></div>';
      const aviso = cacheDe
        ? '<div style="color:#92400e; background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:8px; font-size:12px; margin-bottom:8px;">Sin conexion con la PC ahora mismo: se muestra lo ultimo que se vio (' + new Date(cacheDe).toLocaleTimeString().slice(0, 5) + '). Puede estar desactualizado.</div>'
        : '';
      let porVend = '';
      if ((r.porVendedor || []).length > 1) {
        porVend = '<div class="mp-items" style="margin-top:8px;">' +
          r.porVendedor.map(v => fila(escaparHtml(v.vendedor) + ' (' + v.ventas + ')', 'Ef ' + dineroPOS(v.efectivo) + ' | Tr ' + dineroPOS(v.transferencia))).join('') + '</div>';
      }
      const rep = '<div class="mp-card">' +
        '<div class="mp-top"><span>Reporte de caja</span><span>' + (r.ventas || 0) + (r.ventas === 1 ? ' venta' : ' ventas') + '</span></div>' +
        '<div class="mp-items">' +
          fila('Entro en efectivo', dineroPOS(r.efectivoEntrada)) +
          fila('Entro por transferencia', dineroPOS(r.transferenciaEntrada)) +
          (r.otroEntrada > 0 ? fila('Otros metodos', dineroPOS(r.otroEntrada)) : '') +
          (r.devoluciones > 0
            ? fila('Devuelto en efectivo (' + r.devoluciones + ')', '-' + dineroPOS(r.devolucionEfectivo)) + fila('Devuelto por transferencia', '-' + dineroPOS(r.devolucionTransferencia))
            : '') +
          fila('Neto en efectivo', dineroPOS(r.netoEfectivo), true) +
          fila('Neto por transferencia', dineroPOS(r.netoTransferencia), true) +
          (r.descuentos > 0 ? fila('Descuentos dados', dineroPOS(r.descuentos)) : '') +
        '</div>' + porVend + '</div>';
      const lista = pedidos.length
        ? pedidos.map(renderVentaPOS).join('')
        : '<div style="color:#94a3b8; font-size:13px;">Sin ventas cobradas en este periodo.</div>';
      const mas = d.hayMas ? '<div style="color:#64748b; font-size:12px; margin-top:6px;">Se muestran las 200 ventas mas recientes (el reporte cuenta todas).</div>' : '';
      const pendHtml = pend.length
        ? '<div style="font-size:12px; font-weight:700; color:#92400e; margin:2px 0 6px;">Pendientes de cobro (' + pend.length + ')</div>' + pend.map(renderMiPedido).join('')
        : '';
      cont.innerHTML = aviso + pendHtml + rep + lista + mas;
    }

    function renderVentaPOS(p) {
      const key = p.id + '_' + String(p.hora || '').replace(/\D/g, '');
      posPedidosPorClave[key] = p;
      const abierto = expandidosMP.has(key);
      const devuelto = p.devuelto || {};
      let hayDev = false;
      let quedaPorDevolver = false;
      const lineas = (p.items || []).map(it => {
        const dv = Number(devuelto[it.sku || it.nombre] || 0);
        if (dv > 0) hayDev = true;
        if (Number(it.cantidad) - dv > 0.0001) quedaPorDevolver = true;
        return '<div><span>' + it.cantidad + ' x ' + escaparHtml(it.nombre) + (dv > 0 ? ' (devuelto ' + dv + ')' : '') + '</span><span>$' + (it.precio * it.cantidad).toFixed(2) + '</span></div>';
      }).join('');
      let met = p.metodoPago || 'Efectivo';
      if (met === 'Combinado') met = 'Efectivo ' + dineroPOS(p.pagoEfectivo) + ' + Transferencia ' + dineroPOS((Number(p.pagoTransferencia) || 0) * 2);
      const esHoy = String(p.hora || '').slice(0, 10) === hoyStr();
      const fecha = esHoy ? String(p.hora || '').slice(11, 16) : String(p.hora || '').slice(5, 16);
      const horaTxt = esHoy ? fecha : fecha.replace('-', '/');
      let acciones = '';
      if (esHoy) acciones += '<button class="mp-btn-imprimir" onclick="imprimirMP(' + p.id + ')">Reimprimir</button>';
      if (quedaPorDevolver) acciones += '<button class="mp-btn-editar" onclick="devolverPOS(\'' + key + '\')">Devolver</button>';
      const itemsHtml =
        '<button class="mp-btn-items" id="mp-btnitems-' + key + '" onclick="toggleItemsMP(\'' + key + '\')">' +
          (abierto ? 'Ocultar productos &#9652;' : 'Ver productos &#9662;') + '</button>' +
        '<div class="mp-items" id="mp-items-' + key + '" style="display:' + (abierto ? 'block' : 'none') + ';">' + lineas +
          (p.descuento > 0 ? '<div><span>Descuento aplicado</span><span>-$' + Number(p.descuento).toFixed(2) + '</span></div>' : '') +
        '</div>';
      return '<div class="mp-card">' +
        '<div class="mp-top"><span>Folio #' + p.id + ' - ' + horaTxt + '</span><span class="mp-estado cobrado">' + (hayDev ? (quedaPorDevolver ? 'DEVOL. PARCIAL' : 'DEVUELTO') : 'COBRADO') + '</span></div>' +
        '<div class="mp-total">' + dineroPOS(p.totalCobrado) + ' - ' + escaparHtml(met) + ' - ' + escaparHtml(p.vendedor || '') + '</div>' +
        itemsHtml +
        (acciones ? '<div class="mp-acciones">' + acciones + '</div>' : '') +
      '</div>';
    }

    // Lo que se le devuelve al cliente (vista previa; el servidor lo calcula igual): la venta en
    // transferencia se cobro x2, y una venta combinada se reparte en la misma proporcion que se pago.
    function calcReembolsoPOS(p, elegidos) {
      let base = 0;
      elegidos.forEach(e => {
        const it = (p.items || []).find(i => (i.sku || i.nombre) === (e.sku || e.nombre));
        base += (it ? Number(it.precio) : 0) * e.cantidad;
      });
      const m = p.metodoPago || 'Efectivo';
      let ef = 0, tr = 0;
      if (m === 'Transferencia') { tr = base * 2; }
      else if (m === 'Combinado') {
        const pe = Number(p.pagoEfectivo) || 0, pt = Number(p.pagoTransferencia) || 0;
        if (pe + pt > 0) { ef = base * pe / (pe + pt); tr = base * pt / (pe + pt) * 2; } else { ef = base; }
      } else { ef = base; }
      return { ef: Math.round(ef * 100) / 100, tr: Math.round(tr * 100) / 100 };
    }

    function pedirCantidadDevPOS(x) {
      const e = prompt('Cuantas unidades devuelve de "' + x.it.nombre + '"? (maximo ' + x.max + ', 0 = ninguna):', String(x.max));
      if (e === null) return null;
      const n = parseFloat(String(e).replace(',', '.'));
      if (isNaN(n) || n < 0) { alert('Escribe una cantidad valida.'); return null; }
      if (n > x.max + 0.0001) { alert('De este producto solo se pueden devolver ' + x.max + '.'); return null; }
      return Math.round(n * 100) / 100;
    }

    async function devolverPOS(key) {
      const p = posPedidosPorClave[key];
      if (!p) return;
      const devuelto = p.devuelto || {};
      const lineas = (p.items || [])
        .map(it => ({ it: it, max: Math.round((Number(it.cantidad) - Number(devuelto[it.sku || it.nombre] || 0)) * 100) / 100 }))
        .filter(x => x.max > 0.0001);
      if (!lineas.length) { alert('De esta venta ya se devolvio todo.'); return; }
      let elegidos = [];
      if (lineas.length === 1) {
        const n = pedirCantidadDevPOS(lineas[0]);
        if (n === null) return;
        if (n > 0) elegidos.push({ sku: lineas[0].it.sku, nombre: lineas[0].it.nombre, cantidad: n });
      } else if (confirm('Venta #' + p.id + ': devolver TODO lo que falta por devolver?\n\nAceptar = todo\nCancelar = elegir producto por producto')) {
        elegidos = lineas.map(x => ({ sku: x.it.sku, nombre: x.it.nombre, cantidad: x.max }));
      } else {
        for (const x of lineas) {
          const n = pedirCantidadDevPOS(x);
          if (n === null) return;
          if (n > 0) elegidos.push({ sku: x.it.sku, nombre: x.it.nombre, cantidad: n });
        }
      }
      if (!elegidos.length) { alert('No elegiste ningun producto para devolver.'); return; }
      const motivo = prompt('Motivo de la devolucion (puedes dejarlo vacio):', '');
      if (motivo === null) return;
      const reintegrar = confirm('Regresar estos productos al inventario?\n\nAceptar = si (estan en buen estado)\nCancelar = no (danados, no se vuelven a vender)');
      const rb = calcReembolsoPOS(p, elegidos);
      const partes = [];
      if (rb.ef > 0) partes.push('efectivo ' + dineroPOS(rb.ef));
      if (rb.tr > 0) partes.push('transferencia ' + dineroPOS(rb.tr));
      if (!confirm('Se devuelve al cliente: ' + (partes.join(' + ') || dineroPOS(0)) + (reintegrar ? '' : ' (sin regresar al inventario)') + '.\n\nConfirmas la devolucion?')) return;
      try {
        const res = await fetch('/api/pos/devolucion', { method: 'POST', body: JSON.stringify({ pedidoId: p.id, pedidoHora: p.hora, items: elegidos, motivo: String(motivo).trim(), reintegrarStock: reintegrar, pin: miPin() }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          alert((data && data.error) || 'No se pudo registrar la devolucion.');
          if (data && data.requierePin) abrirLogin((nombreInput.value || '').trim());
          if (data && data.posInactivo) { posLimpiarLocal(); posAplicarUI(true); }
        } else {
          const pa = [];
          if (data.reembolsoEfectivo > 0) pa.push('efectivo ' + dineroPOS(data.reembolsoEfectivo));
          if (data.reembolsoTransferencia > 0) pa.push('transferencia ' + dineroPOS(data.reembolsoTransferencia));
          alert('Devolucion registrada. Devuelve: ' + (pa.join(' + ') || dineroPOS(0)) + '.' +
            (data.impreso ? '' : '\n(No se pudo imprimir el comprobante en la PC' + (data.errorImpresion ? ': ' + data.errorImpresion : '') + ')'));
        }
        cargarMisPedidos();
        cargarCatalogo();
      } catch (e) {
        alert('Sin conexion con la PC: la devolucion NO se registro. Vuelve a intentarlo cuando haya conexion.');
      }
    }

    posAplicarUI(true);
    posVerificar();
    setInterval(posVerificar, 30000);


    // ================= MENSAJES CORTOS A LA CAJA =================
    // La nota opcional viaja con el pedido. Sin productos en el pedido, el mismo boton de enviar
    // manda solo el mensaje. No existe para clientes (autoservicio).
    function actualizarBotonEnviar() {
      const btn = document.getElementById('btnEnviarPedido');
      const nota = document.getElementById('notaPedido');
      if (!btn || btn.disabled) return;
      const soloMensaje = carrito.length === 0 && !!nota && nota.value.trim() !== '' && !autoservicioActivo && puede('mensajes');
      btn.textContent = soloMensaje ? 'Enviar mensaje a la caja' : 'Enviar pedido';
    }

    async function enviarMensajeSolo() {
      const campo = document.getElementById('notaPedido');
      const texto = (campo.value || '').trim();
      const btn = document.getElementById('btnEnviarPedido');
      if (!texto) return;
      if (!(nombreInput.value || '').trim()) { abrirLogin(); return; }
      btn.disabled = true;
      btn.textContent = 'Enviando...';
      try {
        const res = await fetch('/api/mensajes', { method: 'POST', body: JSON.stringify({ vendedor: (nombreInput.value || '').trim(), pin: miPin(), texto: texto }) });
        const data = await res.json().catch(() => null);
        if (!res.ok || !data || !data.ok) {
          mostrarMensaje((data && data.error) || 'No se pudo enviar el mensaje.', false);
          if (data && data.requierePin) abrirLogin((nombreInput.value || '').trim());
        } else {
          campo.value = '';
          mostrarMensaje('Mensaje enviado a la caja.', true);
        }
      } catch (e) {
        mostrarMensaje('Sin conexion con la PC: el mensaje NO se envio (sigue escrito, intentalo de nuevo).', false);
      } finally {
        btn.disabled = false;
        actualizarBotonEnviar();
      }
    }

    const TITULOS_ALERTA = { pedido: 'Pedido modificado por la caja', precio: 'Cambio de precio', stock: 'Cambio de stock', anulado: 'Pedido anulado', permisos: 'Tus permisos cambiaron', mensaje: 'Mensaje de la caja' };

    async function revisarAlertas() {
      const nombre = (nombreInput.value || '').trim();
      if (!nombre) return;
      try {
        const res = await fetch('/api/alertas?vendedor=' + encodeURIComponent(nombre));
        const data = await res.json();
        const alertas = (data && data.alertas) || [];
        for (const a of alertas) {
          beepAsignado();
          try { if (navigator.vibrate) navigator.vibrate([200, 100, 200]); } catch (e) {}
          mostrarMensaje(a.mensaje, false);
          agregarNotificacionCampana(a.mensaje);
          mostrarNotificacionSistema(TITULOS_ALERTA[a.tipo] || 'Aviso de la caja', a.mensaje);
          if (a.tipo === 'permisos') revisarPermisos();
          if (a.tipo === 'pedido' || a.tipo === 'anulado') { try { cargarMisPedidos(true); } catch (e) {} }
          if (!a.tipo) alert(a.mensaje);
        }
      } catch (e) {}
    }
    setInterval(revisarAlertas, 7000);
    setInterval(revisarPedidosAsignados, 5000);
  </script>
</body>
</html>
'@

$htmlEtiquetas = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Etiquetas y códigos de barra - Toto Tools</title>
<style id="pageStyle"></style>
<style>
  * { box-sizing: border-box; font-family: 'Segoe UI', system-ui, Tahoma, Arial, sans-serif; }
  body { margin:0; background:#0f172a; color:#e2e8f0; }
  #app { display:flex; gap:16px; padding:16px; align-items:flex-start; flex-wrap:wrap; }
  .panel { background:#1e293b; border-radius:12px; padding:14px; }
  #ctl { width:390px; max-width:100%; }
  #vista { flex:1; min-width:300px; }
  h1 { font-size:18px; margin:0 0 4px; color:#fff; }
  h2 { font-size:12px; margin:14px 0 6px; color:#94a3b8; text-transform:uppercase; letter-spacing:.5px; }
  input[type=text], input[type=number], select { width:100%; padding:8px; border-radius:6px; border:1px solid #334155; background:#0f172a; color:#e2e8f0; font-size:14px; }
  label.chk { display:flex; align-items:center; gap:8px; font-size:13px; margin:5px 0; }
  .fila { display:flex; gap:8px; align-items:center; }
  .fila > * { flex:1; }
  button { background:#334155; color:#e2e8f0; border:none; padding:8px 12px; border-radius:8px; font-size:13px; font-weight:600; cursor:pointer; }
  button.pri { background:#2563eb; color:#fff; }
  button.pri:hover { background:#1d4ed8; }
  #lista { max-height:290px; overflow:auto; border:1px solid #334155; border-radius:8px; margin-top:8px; }
  .prod { display:flex; align-items:center; gap:8px; padding:6px 8px; border-bottom:1px solid #334155; font-size:13px; }
  .prod .n { flex:1; min-width:0; }
  .prod .n small { display:block; color:#94a3b8; }
  .prod input[type=number] { width:58px; padding:4px 6px; }
  .nota { font-size:12px; color:#94a3b8; margin-top:6px; }
  #hoja { background:#fff; padding:10px; border-radius:8px; min-height:120px; display:flex; flex-wrap:wrap; gap:8px; align-content:flex-start; }
  .et { background:#fff; color:#000; border:1px dashed #94a3b8; overflow:hidden; padding:.8mm 1.2mm; display:flex; flex-direction:column; align-items:center; justify-content:space-between; text-align:center; line-height:1.05; }
  .et .neg { font-weight:700; font-size:calc(var(--k) * 2.3mm); letter-spacing:.3px; }
  .et .nom { font-size:calc(var(--k) * 2.9mm); font-weight:600; max-height:calc(var(--k) * 6.4mm); overflow:hidden; }
  .et .pre { font-size:calc(var(--k) * 6.2mm); font-weight:800; }
  .et .tr { font-size:calc(var(--k) * 2.6mm); }
  .et .bc { width:100%; flex:1; min-height:5mm; display:flex; }
  .et .bc svg { width:100%; height:100%; }
  .et .sk { font-size:calc(var(--k) * 2.3mm); font-family:Consolas, monospace; }
  .vacio { color:#64748b; padding:24px; font-size:14px; }
  @media print {
    body { background:#fff; }
    #ctl, #titulo, .nota-vista { display:none !important; }
    #app { display:block; padding:0; }
    #vista { padding:0; }
    .panel { background:none; padding:0; border-radius:0; }
    #hoja { padding:0; gap:0; border-radius:0; min-height:0; }
    .et { border:none; }
    body.modo-hoja .et { outline:.1mm dashed #aaa; outline-offset:-.1mm; }
    body.modo-rollo .et { page-break-after:always; break-after:page; }
  }
</style>
</head>
<body class="modo-rollo">
<div id="app">
  <div class="panel" id="ctl">
    <h1>Etiquetas y códigos de barra</h1>
    <div class="nota">Elige los productos, cuántas etiquetas de cada uno, y pulsa Imprimir.</div>

    <h2>Productos</h2>
    <input type="text" id="q" placeholder="Buscar por nombre o SKU" oninput="pintarLista()">
    <div class="fila" style="margin-top:8px;">
      <button onclick="marcarVisibles(1)">Marcar visibles</button>
      <button onclick="marcarVisibles(0)">Quitar marcas</button>
      <button onclick="soloBajoStock()" title="Marca los que estan en su minimo de seguridad">Stock bajo</button>
      <button onclick="soloNuevos()" title="Marca los productos que aparecieron por primera vez en el ultimo Excel cargado">Nuevos del Excel</button>
    </div>
    <label class="chk" style="margin-top:8px;"><input type="checkbox" id="oSinStock" onchange="cambioSinStock()"> Ocultar productos sin stock</label>
    <div class="nota" id="notaSinStock" style="margin-top:0;"></div>
    <div id="lista"><div class="vacio">Cargando catálogo...</div></div>
    <div class="nota" id="resumen"></div>

    <h2>Contenido de la etiqueta</h2>
    <label style="font-size:12px; color:#94a3b8;">Nombre del negocio</label>
    <input type="text" id="negocio" value="__NEGOCIO__" oninput="cambio()">
    <label class="chk"><input type="checkbox" id="oNeg" checked onchange="cambio()"> Mostrar nombre del negocio</label>
    <label class="chk"><input type="checkbox" id="oNom" checked onchange="cambio()"> Mostrar nombre del producto</label>
    <label class="chk"><input type="checkbox" id="oPre" checked onchange="cambio()"> Mostrar precio</label>
    <label style="font-size:12px; color:#94a3b8; display:block; margin-top:6px;">Precio que se imprime</label>
    <select id="modoPrecio" onchange="cambio()">
      <option value="efectivo">Solo efectivo</option>
      <option value="transferencia">Solo transferencia (x2)</option>
      <option value="ambos">Ambos (efectivo grande + transferencia pequeño)</option>
    </select>
    <div class="nota">Se queda guardado: todas las etiquetas salen con ese precio hasta que lo cambies.</div>
    <label class="chk"><input type="checkbox" id="oBc" checked onchange="cambio()"> Código de barras (Code 128 con el SKU)</label>
    <label class="chk"><input type="checkbox" id="oSk" checked onchange="cambio()"> Mostrar SKU en texto</label>

    <h2>Estilo de la etiqueta</h2>
    <div class="fila">
      <div><label style="font-size:12px; color:#94a3b8;">Estilo</label>
        <select id="estilo" onchange="cambio()">
          <option value="clasico">Clásico</option><option value="moderno">Moderno</option><option value="banner-top">Banner arriba</option>
          <option value="banner-precio">Banner en el precio</option><option value="marco-doble">Marco doble</option><option value="oferta">Oferta</option>
          <option value="retro">Retro punteado</option><option value="lineas">Líneas</option><option value="minimal">Minimal</option>
          <option value="boutique">Boutique</option><option value="pop">Pop</option><option value="industrial">Industrial</option>
        </select></div>
      <div><label style="font-size:12px; color:#94a3b8;">Letra</label>
        <select id="fuente" onchange="cambio()">
          <option value="arial">Arial</option><option value="verdana">Verdana</option><option value="georgia">Georgia</option>
          <option value="impact">Impact</option><option value="courier">Courier</option>
        </select></div>
    </div>
    <div class="fila" style="margin-top:8px;">
      <div><label style="font-size:12px; color:#94a3b8;">Moneda</label><input type="text" id="moneda" value="$" maxlength="4" oninput="cambio()"></div>
      <div><label style="font-size:12px; color:#94a3b8;">Decimales</label>
        <select id="decimales" onchange="cambio()"><option value="auto">Auto</option><option value="0">Sin decimales</option><option value="2">2 decimales</option></select></div>
    </div>
    <label style="font-size:12px; color:#94a3b8; display:block; margin-top:8px;">Tamaño del precio</label>
    <input type="range" id="tamPrecio" min="50" max="200" value="100" oninput="cambio()">

    <h2>Tamaño y hoja</h2>
    <select id="preset" onchange="aplicarPreset()">
      <option value="57x40">57 x 40 mm (rollo continuo)</option>
      <option value="50x30">50 x 30 mm (rollo)</option>
      <option value="40x25">40 x 25 mm (rollo)</option>
      <option value="58x40">58 x 40 mm (rollo)</option>
      <option value="60x40">60 x 40 mm (rollo)</option>
      <option value="70x37">70 x 37 mm (hoja A4, 3 columnas)</option>
      <option value="custom">Personalizado</option>
    </select>
    <div class="fila" style="margin-top:8px;">
      <div><label style="font-size:12px; color:#94a3b8;">Ancho (mm)</label><input type="number" id="ancho" min="20" max="200" value="50" oninput="cambio(true)"></div>
      <div><label style="font-size:12px; color:#94a3b8;">Alto (mm)</label><input type="number" id="alto" min="12" max="200" value="30" oninput="cambio(true)"></div>
    </div>
    <label style="font-size:12px; color:#94a3b8; display:block; margin-top:8px;">Orientación en el rollo (térmica)</label>
    <select id="orient" onchange="cambio()">
      <option value="h">Horizontal (normal)</option>
      <option value="v90">Vertical (girada 90° a la derecha)</option>
      <option value="v270">Vertical (girada 90° a la izquierda)</option>
    </select>
    <div class="nota">Vertical: el dibujo sale de lado, con el texto corriendo a lo largo del rollo. El "Ancho" es lo que mide a lo largo del papel y el "Alto" lo que mide a lo ancho del rollo (máx. 48 mm en papel de 58 mm).</div>
    <label style="font-size:12px; color:#94a3b8; display:block; margin-top:8px;">Papel</label>
    <select id="modo" onchange="cambio()">
      <option value="rollo">Rollo / etiquetadora (una etiqueta por página)</option>
      <option value="hoja">Hoja A4 con varias etiquetas</option>
    </select>
    <label style="font-size:12px; color:#94a3b8; display:block; margin-top:8px;">Tamaño de letra</label>
    <input type="range" id="escala" min="60" max="160" value="100" oninput="cambio()">

    <h2>Imprimir</h2>
    <div class="fila">
      <div><label style="font-size:12px; color:#94a3b8;">Avance extra entre etiquetas (mm)</label><input type="number" id="feed" min="0" max="30" value="5" oninput="cambio()"></div>
      <label class="chk" style="margin-top:16px;"><input type="checkbox" id="corte" onchange="cambio()"> Cortar después de cada una</label>
    </div>
    <div class="fila" style="margin-top:10px; flex-direction:column; align-items:stretch;">
      <button class="pri" onclick="imprimirTermica()">Imprimir en la térmica (ticket)</button>
      <button onclick="imprimir()">Imprimir en A4 / otra impresora</button>
      <button onclick="location.href='/'">Volver al Panel</button>
    </div>
    <div id="msgTermica" class="nota" style="min-height:16px;"></div>
    <div class="nota">Térmica: sale directo a la impresora del ticket, con el tamaño en mm de arriba (ancho máx. 48 mm en papel de 58 mm). A4: en el cuadro de Chrome usa márgenes "Ninguno" y escala 100 %.</div>
  </div>

  <div class="panel" id="vista">
    <div id="prevRasterNota" style="font-size:13px; color:#94a3b8; margin-bottom:6px;"></div>
    <canvas id="prevRaster" style="display:none; max-width:100%; background:#fff; border-radius:6px; margin-bottom:14px;"></canvas>
    <div id="titulo" style="font-size:13px; color:#94a3b8; margin-bottom:8px;">Vista previa en hoja A4 / navegador <span id="cuenta"></span></div>
    <div id="hoja"><div class="vacio">Marca productos a la izquierda para ver las etiquetas.</div></div>
  </div>
</div>

<script>
  // ---- Code 128 (B, y C para numeros pares) ----
  var C128 = ['212222','222122','222221','121223','121322','131222','122213','122312','132212','221213','221312','231212','112232','122132','122231','113222','123122','123221','223211','221132','221231','213212','223112','312131','311222','321122','321221','312212','322112','322211','212123','212321','232121','111323','131123','131321','112313','132113','132311','211313','231113','231311','112133','112331','132131','113123','113321','133121','313121','211331','231131','213113','213311','213131','311123','311321','331121','312113','312311','332111','314111','221411','431111','111224','111422','121124','121421','141122','141221','112214','112412','122114','122411','142112','142211','241211','221114','413111','241112','134111','111242','121142','121241','114212','124112','124211','411212','421112','421211','212141','214121','412121','111143','111341','131141','114113','114311','411113','411311','113141','114131','311141','411131','211412','211214','211232','2331112'];

  function code128(txt) {
    var v = [];
    if (/^\d{4,}$/.test(txt) && txt.length % 2 === 0) {
      v.push(105);
      for (var i = 0; i < txt.length; i += 2) v.push(parseInt(txt.substr(i, 2), 10));
    } else {
      v.push(104);
      for (var j = 0; j < txt.length; j++) {
        var c = txt.charCodeAt(j);
        if (c < 32 || c > 126) c = 63;
        v.push(c - 32);
      }
    }
    var suma = v[0];
    for (var k = 1; k < v.length; k++) suma += v[k] * k;
    v.push(suma % 103);
    v.push(106);
    return v.map(function (x) { return C128[x]; }).join('');
  }

  function barcodeSVG(txt) {
    var w = code128(txt), x = 10, r = '';
    for (var i = 0; i < w.length; i++) {
      var a = parseInt(w[i], 10);
      if (i % 2 === 0) r += '<rect x="' + x + '" y="0" width="' + a + '" height="50"/>';
      x += a;
    }
    var total = x + 10;
    return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + total + ' 50" preserveAspectRatio="none" shape-rendering="crispEdges" fill="#000">' + r + '</svg>';
  }

  // ---- Estado ----
  var catalogo = [];       // lo que se ve (ya filtrado por "sin stock" si esta marcado)
  var catalogoTodo = [];   // catalogo completo tal como lo manda el servidor
  var sel = {};            // sku|nombre -> copias
  var bajos = {};          // sku -> true
  function clave(p) { return p.sku ? p.sku : ('n:' + p.nombre); }
  function $(id) { return document.getElementById(id); }
  function esc(t) { return String(t == null ? '' : t).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function norm(t) { return String(t || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
  function fmt(n) { return Number(n).toFixed(2).replace(/\.00$/, ''); }

  function guardarAjustes() {
    try {
      localStorage.setItem('etiquetasAjustes', JSON.stringify({
        negocio: $('negocio').value, oNeg: $('oNeg').checked, oNom: $('oNom').checked, oPre: $('oPre').checked,
        modoPrecio: $('modoPrecio').value, estilo: $('estilo').value, fuente: $('fuente').value, tamPrecio: $('tamPrecio').value, moneda: $('moneda').value, decimales: $('decimales').value, oBc: $('oBc').checked, oSk: $('oSk').checked, preset: $('preset').value,
        ancho: $('ancho').value, alto: $('alto').value, modo: $('modo').value, escala: $('escala').value, orient: $('orient').value, feed: $('feed').value, corte: $('corte').checked
      }));
    } catch (e) {}
  }
  function cargarAjustes() {
    try {
      var a = JSON.parse(localStorage.getItem('etiquetasAjustes') || 'null');
      if (!a) return;
      $('negocio').value = a.negocio || $('negocio').value;
      ['oNeg', 'oNom', 'oPre', 'oBc', 'oSk'].forEach(function (id) { $(id).checked = !!a[id]; });
      ['modoPrecio', 'estilo', 'fuente', 'tamPrecio', 'moneda', 'decimales'].forEach(function (id) { if (a[id] != null && a[id] !== '') $(id).value = a[id]; });
      $('preset').value = a.preset || '50x30'; $('ancho').value = a.ancho || 50; $('alto').value = a.alto || 30;
      $('modo').value = a.modo || 'rollo'; $('escala').value = a.escala || 100; $('orient').value = a.orient || 'h';
      $('feed').value = (a.feed == null ? 3 : a.feed); $('corte').checked = !!a.corte;
    } catch (e) {}
  }

  function aplicarPreset() {
    var p = $('preset').value;
    if (p !== 'custom') {
      var m = p.split('x');
      $('ancho').value = m[0]; $('alto').value = m[1];
      if (p === '70x37') $('modo').value = 'hoja';
    }
    cambio();
  }

  function cambio(manual) {
    if (manual === true) $('preset').value = 'custom';
    document.body.className = ($('modo').value === 'hoja') ? 'modo-hoja' : 'modo-rollo';
    var w = Math.max(20, +$('ancho').value || 50), h = Math.max(12, +$('alto').value || 30);
    $('pageStyle').textContent = ($('modo').value === 'hoja')
      ? '@page { size: A4; margin: 8mm; }'
      : '@page { size: ' + w + 'mm ' + h + 'mm; margin: 0; }';
    guardarAjustes();
    pintarVista();
  }

  function hayControlStockEt() {
    return catalogoTodo.some(function (p) { return p.stock !== null && p.stock !== undefined; });
  }
  function aplicarFiltroStock() {
    var ocultar = $('oSinStock').checked && hayControlStockEt();
    catalogo = ocultar
      ? catalogoTodo.filter(function (p) { return !(p.stock === null || p.stock === undefined || p.stock <= 0); })
      : catalogoTodo;
    $('notaSinStock').textContent = ocultar
      ? ((catalogoTodo.length - catalogo.length) + ' producto(s) sin stock ocultos.')
      : (hayControlStockEt() ? '' : 'El Excel no trae columna de cantidad: no se puede saber qué está sin stock.');
  }
  function cambioSinStock() {
    aplicarFiltroStock();
    // Los que quedaron ocultos dejan de estar marcados (no se imprimen a escondidas).
    var visibles = {};
    catalogo.forEach(function (p) { visibles[clave(p)] = true; });
    Object.keys(sel).forEach(function (k) { if (!visibles[k]) delete sel[k]; });
    pintarLista(); pintarVista();
  }

  function pintarLista() {
    var q = norm($('q').value);
    var visibles = catalogo.filter(function (p) { return !q || norm(p.nombre).indexOf(q) >= 0 || norm(p.sku).indexOf(q) >= 0; }).slice(0, 300);
    $('lista').innerHTML = visibles.length ? visibles.map(function (p) {
      var k = clave(p), on = sel[k] > 0;
      return '<div class="prod"><input type="checkbox" ' + (on ? 'checked' : '') + ' onchange="marcar(' + esc(JSON.stringify(k)) + ',this.checked)">' +
        '<div class="n">' + esc(p.nombre) + '<small>' + esc(p.sku || 'sin SKU') + ' · $' + fmt(p.precio) + '</small></div>' +
        '<input type="number" min="1" max="500" value="' + (on ? sel[k] : 1) + '" ' + (on ? '' : 'disabled') + ' onchange="copias(' + esc(JSON.stringify(k)) + ',this.value)"></div>';
    }).join('') : '<div class="vacio">Sin resultados.</div>';
    resumen();
  }
  function marcar(k, on) { if (on) sel[k] = sel[k] > 0 ? sel[k] : 1; else delete sel[k]; pintarLista(); pintarVista(); }
  function copias(k, v) { var n = Math.max(1, Math.min(500, parseInt(v, 10) || 1)); sel[k] = n; pintarVista(); resumen(); }
  function marcarVisibles(on) {
    var q = norm($('q').value);
    catalogo.filter(function (p) { return !q || norm(p.nombre).indexOf(q) >= 0 || norm(p.sku).indexOf(q) >= 0; }).slice(0, 300).forEach(function (p) {
      if (on) sel[clave(p)] = sel[clave(p)] > 0 ? sel[clave(p)] : 1; else delete sel[clave(p)];
    });
    pintarLista(); pintarVista();
  }
  function soloBajoStock() {
    sel = {};
    catalogo.forEach(function (p) { if (p.sku && bajos[p.sku]) sel[clave(p)] = 1; });
    pintarLista(); pintarVista();
  }
  var nuevosSet = {}, nuevosFecha = '';
  function soloNuevos() {
    var total = Object.keys(nuevosSet).length;
    if (!total) { alert('El ultimo Excel no trajo productos nuevos (o todavia no se ha cargado ninguno despues de activar esta funcion).'); return; }
    sel = {};
    var n = 0;
    catalogo.forEach(function (p) { if (nuevosSet[clave(p)]) { sel[clave(p)] = 1; n++; } });
    pintarLista(); pintarVista();
    if (n < total) alert('Se marcaron ' + n + ' de ' + total + ' productos nuevos; el resto esta oculto por "Ocultar productos sin stock".');
  }
  function resumen() {
    var ks = Object.keys(sel), total = 0;
    ks.forEach(function (k) { total += sel[k]; });
    $('resumen').textContent = ks.length + ' producto(s) marcados · ' + total + ' etiqueta(s)';
  }

  function etiquetaHTML(p) {
    var neg = $('negocio').value.trim();
    var s = '<div class="et" style="width:' + $('ancho').value + 'mm;height:' + $('alto').value + 'mm;--k:' + (($('escala').value || 100) / 100) + '">';
    if ($('oNeg').checked && neg) s += '<div class="neg">' + esc(neg) + '</div>';
    if ($('oNom').checked) s += '<div class="nom">' + esc(p.nombre) + '</div>';
    var modo = $('modoPrecio').value, mon = $('moneda').value || '$';
    if ($('oPre').checked) {
      s += '<div class="pre">' + mon + fmt(modo === 'transferencia' ? p.precio * 2 : p.precio) + '</div>';
      if (modo === 'ambos') s += '<div class="tr">Transferencia: ' + mon + fmt(p.precio * 2) + '</div>';
    }
    if ($('oBc').checked && p.sku) s += '<div class="bc">' + barcodeSVG(String(p.sku)) + '</div>';
    if ($('oSk').checked && p.sku) s += '<div class="sk">' + esc(p.sku) + '</div>';
    return s + '</div>';
  }

  function pintarVista() {
    var out = [], n = 0;
    catalogo.forEach(function (p) {
      var c = sel[clave(p)] || 0;
      for (var i = 0; i < c; i++) { out.push(etiquetaHTML(p)); n++; }
    });
    $('hoja').innerHTML = out.length ? out.join('') : '<div class="vacio">Marca productos a la izquierda para ver las etiquetas.</div>';
    $('cuenta').textContent = n ? '(' + n + ' etiquetas)' : '';
    resumen();
    pintarRaster();
  }

  function imprimir() {
    if (!Object.keys(sel).length) { alert('Marca al menos un producto.'); return; }
    window.print();
  }


  // ================= Impresion por raster (mismo dibujo que tu app de etiquetas) =================
  var DOTS_MM = 8, BANDA = 64;
  var cfg = {};
  function clamp(v, a, b) { return Math.min(b, Math.max(a, v)); }
  function estiloActual() { return ESTILOS[cfg.estilo] || ESTILOS.clasico; }
  function precioEtiquetaR(it) {
    var n = parsePrecio(it.precio);
    if (isNaN(n)) return it.precio;
    return cfg.modoPrecio === 'transferencia' ? n * 2 : n;
  }
  function leerCfgRaster() {
    cfg = {
      ancho: clamp(parseFloat($('ancho').value) || 57, 20, 120),
      alto: clamp(parseFloat($('alto').value) || 40, 15, 200),
      empresa: $('negocio').value, moneda: $('moneda').value, decimales: $('decimales').value,
      verEmpresa: $('oNeg').checked, verPrecio: $('oPre').checked, verBarras: $('oBc').checked, verCodigo: $('oSk').checked,
      modoPrecio: $('modoPrecio').value, tamPrecio: clamp(parseInt($('tamPrecio').value, 10) || 100, 50, 200),
      estilo: $('estilo').value, fuente: $('fuente').value, orient: $('orient').value
    };
    return cfg;
  }
  function itemsMarcados() {
    var out = [];
    catalogo.forEach(function (p) {
      var c = sel[clave(p)] || 0;
      if (c > 0) out.push({ nombre: p.nombre, precio: p.precio, codigo: p.sku || '', cant: c });
    });
    return out;
  }

const FUENTES = {
    arial:      'Arial, Helvetica, sans-serif',
    verdana:    'Verdana, Geneva, sans-serif',
    georgia:    'Georgia, "Times New Roman", serif',
    montserrat: '"Montserrat", Arial, sans-serif',
    condensed:  '"Roboto Condensed", "Arial Narrow", sans-serif',
    playfair:   '"Playfair Display", Georgia, serif',
    cinzel:     '"Cinzel", Georgia, serif',
    impact:     'Impact, "Arial Narrow", sans-serif',
    courier:    '"Courier Prime", "Courier New", monospace'
  };

const ESTILOS = {
    clasico:        { padding: 1.5, marco: 'simple',   header: 'linea',   precioBanner: false, alinear: 'centro',    factorNombre: 0.13, factorPrecio: 0.26 },
    moderno:        { padding: 1.3, marco: 'ninguno',  header: 'ninguno', precioBanner: false, alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.27 },
    'banner-top':   { padding: 1.5, marco: 'ninguno',  header: 'banner',  precioBanner: false, alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.27 },
    'banner-precio':{ padding: 1.4, marco: 'ninguno',  header: 'linea',   precioBanner: true,  alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.24 },
    'marco-doble':  { padding: 2.0, marco: 'doble',    header: 'linea',   precioBanner: false, alinear: 'centro',    factorNombre: 0.13, factorPrecio: 0.25 },
    oferta:         { padding: 1.5, marco: 'ninguno',  header: 'banner',  precioBanner: false, alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.27 },
    retro:          { padding: 1.7, marco: 'punteado', header: 'punteada',precioBanner: false, alinear: 'centro',    factorNombre: 0.13, factorPrecio: 0.25 },
    lineas:         { padding: 1.4, marco: 'ninguno',  header: 'lineas',  precioBanner: false, alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.26 },
    minimal:        { padding: 1.2, marco: 'ninguno',  header: 'ninguno', precioBanner: false, alinear: 'izquierda', factorNombre: 0.12, factorPrecio: 0.28 },
    boutique:       { padding: 2.1, marco: 'simple',   header: 'linea',   precioBanner: false, alinear: 'centro',    factorNombre: 0.12, factorPrecio: 0.24 },
    pop:            { padding: 1.7, marco: 'grueso',   header: 'linea',   precioBanner: false, alinear: 'centro',    factorNombre: 0.13, factorPrecio: 0.29 },
    industrial:     { padding: 1.6, marco: 'grueso',   header: 'banner',  precioBanner: false, alinear: 'centro',    factorNombre: 0.13, factorPrecio: 0.27 }
  };

function parsePrecio(txt) {
    let t = String(txt ?? '').replace(/[^\d.,-]/g, '');
    if (!t) return NaN;
    const c = t.lastIndexOf(','), p = t.lastIndexOf('.');
    if (c > -1 && p > -1) {
      // el último separador es el decimal
      t = c > p ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
    } else if (c > -1) {
      t = t.replace(',', '.');
    }
    return parseFloat(t);
  }

function formatoPrecio(txt) {
    const n = parsePrecio(txt);
    const mon = cfg.moneda.trim();
    let cuerpo;
    if (isNaN(n)) {
      cuerpo = String(txt ?? '').trim() || '0';
    } else if (cfg.decimales === '0') {
      cuerpo = String(Math.round(n));
    } else if (cfg.decimales === '2') {
      cuerpo = n.toFixed(2);
    } else {
      cuerpo = Number.isInteger(n) ? String(n) : n.toFixed(2);
    }
    if (!mon) return cuerpo;
    return mon.length <= 2 ? mon + cuerpo : cuerpo + ' ' + mon;
  }

function code128Modulos(texto) {
    const limpio = String(texto).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^\x20-\x7e]/g, '');
    if (!limpio) return null;
    const valores = [104];
    let suma = 104;
    for (let i = 0; i < limpio.length; i++) {
      const v = limpio.charCodeAt(i) - 32;
      valores.push(v);
      suma += v * (i + 1);
    }
    valores.push(suma % 103, 106);
    const barras = [];
    let x = 10;
    valores.forEach(v => {
      let negro = true;
      for (const d of C128[v]) {
        const w = +d;
        if (negro) barras.push([x, w]);
        x += w;
        negro = !negro;
      }
    });
    return { barras, total: x + 10 };
  }

function partirLineas(g, texto, maxW) {
    const palabras = texto.split(/\s+/).filter(Boolean);
    const lineas = [];
    let actual = '';
    for (const pal of palabras) {
      const prueba = actual ? actual + ' ' + pal : pal;
      if (g.measureText(prueba).width <= maxW) { actual = prueba; continue; }
      if (actual) lineas.push(actual);
      if (g.measureText(pal).width <= maxW) { actual = pal; continue; }
      let trozo = '';
      for (const ch of pal) {
        if (trozo && g.measureText(trozo + ch).width > maxW) { lineas.push(trozo); trozo = ch; }
        else trozo += ch;
      }
      actual = trozo;
    }
    if (actual) lineas.push(actual);
    return lineas;
  }

function trazarRect(g, x, y, w, h, grosor) {
    g.fillRect(x, y, w, grosor);
    g.fillRect(x, y + h - grosor, w, grosor);
    g.fillRect(x, y, grosor, h);
    g.fillRect(x + w - grosor, y, grosor, h);
  }

function trazarRectPunteado(g, x, y, w, h, grosor) {
    const paso = grosor * 3, trazo = paso * 0.6;
    for (let px = x; px < x + w; px += paso) {
      const t = Math.min(trazo, x + w - px);
      g.fillRect(px, y, t, grosor);
      g.fillRect(px, y + h - grosor, t, grosor);
    }
    for (let py = y; py < y + h; py += paso) {
      const t = Math.min(trazo, y + h - py);
      g.fillRect(x, py, grosor, t);
      g.fillRect(x + w - grosor, py, grosor, t);
    }
  }

function trazarLineaPunteada(g, x, y, w, grosor) {
    const paso = grosor * 4, trazo = paso * 0.6;
    for (let px = x; px < x + w; px += paso) g.fillRect(px, y, Math.min(trazo, x + w - px), grosor);
  }

function rasterEtiqueta(it) {
    const D = DOTS_MM, wl = cfg.ancho, hl = cfg.alto;
    const est = estiloActual();
    const famBase = FUENTES[cfg.fuente] || FUENTES.arial;
    // Vertical: la etiqueta se dibuja como se lee (ancho x alto) y al final se gira 90 grados,
    // asi sale de lado en el rollo. Lo que mide a lo ancho del papel es entonces el "alto".
    const rot = cfg.orient === 'v90' || cfg.orient === 'v270';
    const maxPapel = (rot ? hl : wl) <= 65 ? 384 : 576;
    const W = rot ? Math.max(8, Math.round(wl * D)) : Math.max(8, Math.floor(Math.min(wl * D, maxPapel) / 8) * 8);
    const H = rot ? Math.max(8, Math.floor(Math.min(hl * D, maxPapel) / 8) * 8) : Math.round(hl * D);
    const cv = document.createElement('canvas');
    cv.width = W; cv.height = H;
    const g = cv.getContext('2d', { willReadFrequently: true });
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    g.fillStyle = '#000'; g.textBaseline = 'top';

    // Marco: dibuja el borde según el estilo y calcula cuánto espacio ocupa
    let inset = 0;
    if (est.marco === 'simple' || est.marco === 'grueso') {
      const grosor = Math.max(1, Math.round((est.marco === 'grueso' ? 0.6 : 0.3) * D));
      trazarRect(g, 0, 0, W, H, grosor);
      inset = grosor + Math.round(0.5 * D);
    } else if (est.marco === 'punteado') {
      const grosor = Math.max(1, Math.round(0.35 * D));
      trazarRectPunteado(g, 0, 0, W, H, grosor);
      inset = grosor + Math.round(0.5 * D);
    } else if (est.marco === 'doble') {
      const grosor = Math.max(1, Math.round(0.22 * D)), hueco = Math.round(0.9 * D);
      trazarRect(g, 0, 0, W, H, grosor);
      trazarRect(g, hueco, hueco, W - 2 * hueco, H - 2 * hueco, grosor);
      inset = hueco + grosor + Math.round(0.3 * D);
    }

    const pad = Math.round(est.padding * D) + inset, inner = W - 2 * pad;
    const izq = est.alinear === 'izquierda';
    g.textAlign = izq ? 'left' : 'center';
    const cx = izq ? pad : W / 2;

    const fuente = (peso, px, fam) => `${peso} ${px}px ${fam || famBase}`;
    const ajustar1 = (txt, peso, px, fam, min) => {
      while (px > min) {
        g.font = fuente(peso, px, fam);
        if (g.measureText(txt).width <= inner) break;
        px -= 1;
      }
      g.font = fuente(peso, px, fam);
      return px;
    };

    // Encabezado
    let y = pad;
    if (cfg.verEmpresa && cfg.empresa.trim()) {
      const txt = cfg.empresa.trim().toUpperCase();
      if (est.header === 'banner') {
        const fs = ajustar1(txt, 'bold', clamp(hl * 0.075, 2.2, 3.5) * D, famBase, 8);
        const altoBanner = Math.round(fs * 1.7);
        g.fillRect(0, y - Math.round(0.4 * D), W, altoBanner);
        g.fillStyle = '#fff';
        g.fillText(txt, cx, y + Math.round((altoBanner - fs) / 2) - Math.round(0.4 * D));
        g.fillStyle = '#000';
        y += altoBanner + 4;
      } else if (est.header === 'lineas') {
        g.fillRect(pad, y, inner, 2); y += 5;
        const fs = ajustar1(txt, 'bold', clamp(hl * 0.075, 2.2, 3.5) * D, famBase, 8);
        g.fillText(txt, cx, y);
        y += Math.round(fs * 1.15) + 3;
        g.fillRect(pad, y, inner, 2); y += 7;
      } else {
        const fs = ajustar1(txt, 'bold', clamp(hl * 0.075, 2.2, 3.5) * D, famBase, 8);
        g.fillText(txt, cx, y);
        y += Math.round(fs * 1.15) + 3;
        if (est.header === 'linea') { g.fillRect(pad, y, inner, 2); y += 7; }
        else if (est.header === 'punteada') { trazarLineaPunteada(g, pad, y, inner, 2); y += 7; }
        else { y += 3; }
      }
    }

    // Parte inferior (de abajo hacia arriba): código en texto, barras, precio
    let yb = H - pad;
    const conBarras = cfg.verBarras && it.codigo, conCodigo = cfg.verCodigo && it.codigo;
    if (conCodigo) {
      const fs = ajustar1(it.codigo, 'normal', 2.4 * D, 'Consolas, "Courier New", monospace', 9);
      yb -= Math.round(fs * 1.1);
      g.fillText(it.codigo, cx, yb);
    }
    if (conBarras) {
      const m = code128Modulos(it.codigo);
      if (m) {
        const hb = Math.round(clamp(hl * 0.2, 5, 10) * D);
        const mod = Math.max(1, Math.min(3, Math.floor(inner / m.total)));
        const x0 = izq ? pad : Math.round((W - m.total * mod) / 2);
        yb -= hb;
        m.barras.forEach(([x, w]) => g.fillRect(x0 + x * mod, yb, w * mod, hb));
        yb -= 2;
      }
    }
    if (cfg.verPrecio) {
      const ptxt = formatoPrecio(precioEtiquetaR(it));
      const ps = ajustar1(ptxt, 'bold', clamp(hl * est.factorPrecio, 5, 14) * D * (cfg.tamPrecio / 100), famBase, 20);
      if (est.precioBanner) {
        const altoBanner = Math.round(ps * 1.35);
        yb -= altoBanner;
        g.fillRect(pad, yb, inner, altoBanner);
        g.fillStyle = '#fff';
        g.fillText(ptxt, cx, yb + Math.round((altoBanner - ps) / 2));
        g.fillStyle = '#000';
        yb -= 3;
      } else {
        yb -= Math.round(ps * 1.05);
        g.fillText(ptxt, cx, yb);
      }
    }

    if (cfg.verPrecio && cfg.modoPrecio === 'ambos') {
      const st = 'Transf. ' + formatoPrecio(parsePrecio(it.precio) * 2);
      const sf = ajustar1(st, 'normal', 2.6 * D, famBase, 8);
      yb -= Math.round(sf * 1.15);
      g.fillText(st, cx, yb);
    }

    // Nombre en el espacio que queda
    const aH = yb - y - 4;
    if (aH > 8) {
      const fNombre = cfg.verPrecio ? est.factorNombre : est.factorNombre * 1.8;
      let fs = clamp(hl * fNombre, 3, cfg.verPrecio ? 6 : 10) * D, lineas = [];
      const nombre = it.nombre || 'Producto';
      for (;;) {
        g.font = fuente('bold', fs, famBase);
        lineas = partirLineas(g, nombre, inner);
        if (lineas.length * fs * 1.1 <= aH || fs <= 12) break;
        fs -= 1;
      }
      lineas = lineas.slice(0, Math.max(1, Math.floor(aH / (fs * 1.1))));
      let ty = y + Math.max(0, (aH - lineas.length * fs * 1.1) / 2);
      lineas.forEach(l => { g.fillText(l, cx, ty); ty += fs * 1.1; });
    }

    // Si es vertical, se gira el dibujo 90 grados antes de pasarlo a puntos
    let cvF = cv, gF = g, WF = W, HF = H;
    if (rot) {
      cvF = document.createElement('canvas'); cvF.width = H; cvF.height = W;
      gF = cvF.getContext('2d', { willReadFrequently: true });
      gF.fillStyle = '#fff'; gF.fillRect(0, 0, H, W);
      if (cfg.orient === 'v90') { gF.translate(H, 0); gF.rotate(Math.PI / 2); }
      else { gF.translate(0, W); gF.rotate(-Math.PI / 2); }
      gF.drawImage(cv, 0, 0);
      WF = H; HF = W;
    }

    // A 1 bit (1 = negro)
    const img = gF.getImageData(0, 0, WF, HF).data;
    const bpr = WF / 8;
    const bits = new Uint8Array(bpr * HF);
    for (let py = 0; py < HF; py++) {
      for (let px = 0; px < WF; px++) {
        const i = (py * WF + px) * 4;
        if (img[i] * 0.299 + img[i + 1] * 0.587 + img[i + 2] * 0.114 < 150) bits[py * bpr + (px >> 3)] |= 0x80 >> (px & 7);
      }
    }
    return { W: WF, H: HF, bpr, bits, cv: cvF };
  }

  function pintarRaster() {
    var cv = $('prevRaster'), nota = $('prevRasterNota');
    if (!cv) return;
    var lista = itemsMarcados();
    if (!lista.length) { cv.style.display = 'none'; nota.textContent = 'Marca un producto para ver cómo sale en la térmica.'; return; }
    try {
      leerCfgRaster();
      var r = rasterEtiqueta(lista[0]);
      cv.width = r.W; cv.height = r.H; cv.getContext('2d').drawImage(r.cv, 0, 0);
      cv.style.display = 'block';
      nota.textContent = 'Así sale en la térmica (primer producto marcado, ' + cfg.ancho + ' x ' + cfg.alto + ' mm' + (cfg.orient === 'h' ? ', horizontal' : ', vertical') + ').';
    } catch (e) { cv.style.display = 'none'; nota.textContent = 'No se pudo dibujar la vista previa.'; }
  }

  function bytesEtiquetaR(it) {
    var r = rasterEtiqueta(it), partes = [];
    for (var y = 0; y < r.H; y += BANDA) {
      var n = Math.min(BANDA, r.H - y);
      partes.push(Uint8Array.of(0x1d, 0x76, 0x30, 0x00, r.bpr & 255, r.bpr >> 8, n & 255, n >> 8));
      partes.push(r.bits.subarray(y * r.bpr, (y + n) * r.bpr));
    }
    var mm = parseFloat($('feed').value); if (isNaN(mm)) mm = 5;
    partes.push(Uint8Array.of(0x1b, 0x4a, clamp(Math.round(mm * DOTS_MM), 0, 255)));
    if ($('corte').checked) partes.push(Uint8Array.of(0x1d, 0x56, 0x42, 3));
    var total = partes.reduce(function (s, p) { return s + p.length; }, 0), out = new Uint8Array(total), o = 0;
    partes.forEach(function (p) { out.set(p, o); o += p.length; });
    return out;
  }

  var imprimiendoTermica = false;
  async function imprimirTermica() {
    var m = $('msgTermica');
    var lista = itemsMarcados();
    if (!lista.length) { m.style.color = '#fca5a5'; m.textContent = 'Marca al menos un producto.'; return; }
    if (imprimiendoTermica) { m.textContent = 'Todavía se está imprimiendo...'; return; }
    imprimiendoTermica = true;
    var total = 0; lista.forEach(function (i) { total += i.cant; });
    var hechas = 0, lote = [], loteN = 0;
    async function enviarLote() {
      if (!lote.length) return;
      var len = lote.reduce(function (s, p) { return s + p.length; }, 0), buf = new Uint8Array(len + 2), o = 2;
      buf[0] = 0x1b; buf[1] = 0x40;
      lote.forEach(function (p) { buf.set(p, o); o += p.length; });
      var r = await fetch('/api/etiquetas/raw', { method: 'POST', body: buf, headers: { 'Content-Type': 'application/octet-stream' } });
      var d = await r.json();
      if (!d.ok) throw new Error(d.error || 'La impresora no respondio.');
      lote = []; loteN = 0;
    }
    try {
      leerCfgRaster();
      for (var a = 0; a < lista.length; a++) {
        var datos = bytesEtiquetaR(lista[a]);
        for (var n = 0; n < lista[a].cant; n++) {
          lote.push(datos); loteN++; hechas++;
          m.style.color = '#94a3b8'; m.textContent = 'Preparando ' + hechas + ' de ' + total + '...';
          if (loteN >= 20) await enviarLote();
        }
      }
      await enviarLote();
      m.style.color = '#86efac'; m.textContent = 'Listo: ' + total + ' etiqueta(s) enviada(s) a la impresora.';
    } catch (e) {
      m.style.color = '#fca5a5'; m.textContent = 'No se pudo imprimir: ' + (e.message || e);
    }
    imprimiendoTermica = false;
  }

  async function iniciar() {
    cargarAjustes();
    document.body.className = ($('modo').value === 'hoja') ? 'modo-hoja' : 'modo-rollo';
    try {
      var r = await fetch('/api/catalogo'); catalogoTodo = await r.json();
      if (!Array.isArray(catalogoTodo)) catalogoTodo = [];
    } catch (e) { catalogoTodo = []; }
    // Por defecto sigue el ajuste del panel "Ocultar sin stock a vendedores".
    try {
      var rc = await fetch('/api/config'); var cc = await rc.json();
      $('oSinStock').checked = !!(cc && cc.ocultarSinStock);
    } catch (e) {}
    aplicarFiltroStock();
    try {
      var r2 = await fetch('/api/reabastecer'); var d2 = await r2.json();
      (d2.items || []).forEach(function (i) { if (i.sku) bajos[i.sku] = true; });
    } catch (e) {}
    try {
      var r3 = await fetch('/api/catalogo/nuevos'); var d3 = await r3.json();
      nuevosFecha = d3.fecha || '';
      (d3.skus || []).forEach(function (k) { nuevosSet[k] = true; });
    } catch (e) {}
    cambio();
    pintarLista();
  }
  iniciar();
</script>
</body>
</html>
'@
$htmlEtiquetas = $htmlEtiquetas.Replace('__NEGOCIO__', [System.Net.WebUtility]::HtmlEncode([string]$reciboNombre))

$htmlMetricas = @'
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Metricas - Toto Tools</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; font-family: 'Segoe UI', system-ui, Tahoma, Arial, sans-serif; }
  body { background:#0f172a; color:#e2e8f0; padding:18px; }
  h1 { font-size:19px; margin-bottom:14px; }
  .tabs { display:flex; gap:8px; margin-bottom:18px; flex-wrap:wrap; }
  .tabs button { background:#1e293b; color:#cbd5e1; border:none; padding:9px 16px; border-radius:20px; font-size:13px; font-weight:600; cursor:pointer; }
  .tabs button.activo { background:#2563eb; color:#fff; }
  .resumen { display:flex; gap:14px; flex-wrap:wrap; margin-bottom:20px; }
  .tarjeta { background:#1e293b; border-radius:12px; padding:16px 20px; min-width:150px; }
  .tarjeta .num { font-size:26px; font-weight:800; }
  .tarjeta .lbl { font-size:12px; color:#94a3b8; margin-top:2px; }
  .panel { background:#1e293b; border-radius:12px; padding:16px 18px; margin-bottom:18px; }
  .panel h2 { font-size:14px; text-transform:uppercase; letter-spacing:0.5px; color:#94a3b8; margin-bottom:12px; }
  .fila { display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:13px; }
  .fila .nombre { width:150px; flex-shrink:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .fila .barra-fondo { flex:1; background:#0f172a; border-radius:6px; height:20px; overflow:hidden; }
  .fila .barra { background:#2563eb; height:100%; border-radius:6px; min-width:2px; }
  .fila .valor { width:90px; text-align:right; flex-shrink:0; font-weight:600; }
  .fila.hora .nombre { width:34px; font-size:11px; color:#94a3b8; }
  .fila.hora .barra { background:#7c3aed; }
  .vacio { color:#64748b; font-size:13px; }
  #cargando { color:#94a3b8; font-size:13px; }
</style>
</head>
<body>
  <h1>Metricas - Toto Tools</h1>
  <div class="tabs">
    <button id="tabHoy" onclick="cambiarRango(1)">Hoy</button>
    <button id="tabSemana" onclick="cambiarRango(7)">Ultimos 7 dias</button>
    <button id="tabMes" onclick="cambiarRango(30)">Ultimos 30 dias</button>
  </div>
  <div id="cargando">Cargando...</div>
  <div id="contenido" style="display:none;">
    <div class="resumen">
      <div class="tarjeta"><div class="num" id="numPedidos">0</div><div class="lbl">Pedidos cobrados</div></div>
      <div class="tarjeta"><div class="num" id="numFacturado">$0.00</div><div class="lbl">Total facturado</div></div>
      <div class="tarjeta"><div class="num" id="numHoraPico">-</div><div class="lbl">Hora pico</div></div>
    </div>
    <div class="panel">
      <h2>Productos mas vendidos</h2>
      <div id="listaProductos"></div>
    </div>
    <div class="panel">
      <h2>Facturado por vendedor</h2>
      <div id="listaVendedores"></div>
    </div>
    <div class="panel">
      <h2>Pedidos por hora del dia</h2>
      <div id="listaHoras"></div>
    </div>
  </div>

  <script>
    let rangoActual = 1;
    async function cambiarRango(dias) {
      rangoActual = dias;
      document.getElementById('tabHoy').classList.toggle('activo', dias === 1);
      document.getElementById('tabSemana').classList.toggle('activo', dias === 7);
      document.getElementById('tabMes').classList.toggle('activo', dias === 30);
      await cargar();
    }
    function barra(valor, max, texto) {
      const pct = max > 0 ? Math.max(2, Math.round((valor / max) * 100)) : 0;
      return '<div class="barra-fondo"><div class="barra" style="width:' + pct + '%;"></div></div><div class="valor">' + texto + '</div>';
    }
    async function cargar() {
      document.getElementById('cargando').style.display = 'block';
      document.getElementById('contenido').style.display = 'none';
      try {
        const res = await fetch('/api/metricas?dias=' + rangoActual);
        const m = await res.json();
        document.getElementById('numPedidos').textContent = m.pedidosCobrados;
        document.getElementById('numFacturado').textContent = '$' + Number(m.totalFacturado).toFixed(2);
        document.getElementById('numHoraPico').textContent = (m.horaPico === null || m.horaPico === undefined) ? '-' : (String(m.horaPico).padStart(2, '0') + ':00');

        const maxProd = Math.max(1, ...m.topProductos.map(p => p.cantidad));
        document.getElementById('listaProductos').innerHTML = m.topProductos.length === 0 ? '<div class="vacio">Sin ventas en este periodo.</div>' :
          m.topProductos.map(p => '<div class="fila"><div class="nombre" title="' + p.nombre + '">' + p.nombre + '</div>' + barra(p.cantidad, maxProd, p.cantidad + ' uds') + '</div>').join('');

        const maxVend = Math.max(1, ...m.porVendedor.map(v => v.monto));
        document.getElementById('listaVendedores').innerHTML = m.porVendedor.length === 0 ? '<div class="vacio">Sin ventas en este periodo.</div>' :
          m.porVendedor.map(v => '<div class="fila"><div class="nombre" title="' + v.vendedor + '">' + v.vendedor + '</div>' + barra(v.monto, maxVend, '$' + v.monto.toFixed(2)) + '</div>').join('');

        const maxHora = Math.max(1, ...m.porHora.map(h => h.pedidos));
        document.getElementById('listaHoras').innerHTML = m.porHora.map(h =>
          '<div class="fila hora"><div class="nombre">' + String(h.hora).padStart(2, '0') + 'h</div>' + barra(h.pedidos, maxHora, String(h.pedidos)) + '</div>'
        ).join('');

        document.getElementById('cargando').style.display = 'none';
        document.getElementById('contenido').style.display = 'block';
      } catch (e) {
        document.getElementById('cargando').textContent = 'No se pudieron cargar las metricas.';
      }
    }
    cambiarRango(1);
  </script>
</body>
</html>
'@

# ------------------------------------------------------------------
# LICENCIA: pagina de bloqueo y script guardian para las pantallas
# ------------------------------------------------------------------
$global:htmlSinLicencia = @'
<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Licencia</title></head>
<body style="margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#111;color:#fff;font-family:Arial,sans-serif;text-align:center;padding:24px;box-sizing:border-box">
<div><div style="font-size:56px">&#128274;</div><h2>Licencia vencida o no v&aacute;lida</h2>
<p style="color:#bbb">El sistema est&aacute; bloqueado. Contacta al proveedor para renovar la licencia.</p></div></body></html>
'@
$global:licGuardJs = @'
<script>(function(){var of=window.fetch;function bloquear(){try{localStorage.setItem('tt_lic_bloq','1')}catch(e){}if(document.getElementById('tt-lic-block'))return;var d=document.createElement('div');d.id='tt-lic-block';d.style.cssText='position:fixed;left:0;top:0;right:0;bottom:0;z-index:2147483647;background:#111;color:#fff;font-family:Arial,sans-serif;text-align:center;display:flex;align-items:center;justify-content:center;padding:24px';d.innerHTML='<div><div style="font-size:56px">&#128274;</div><h2>Licencia vencida</h2><p style="color:#bbb">Contacta al proveedor para renovarla.</p></div>';(document.body||document.documentElement).appendChild(d);}
window.fetch=function(){return of.apply(this,arguments).then(function(r){if(r&&r.status===402)bloquear();return r;});};
function aviso(){of('/api/licencia',{cache:'no-store'}).then(function(r){if(r.status===402){bloquear();return null;}return r.json();}).then(function(j){if(!j||!j.ok)return;try{localStorage.removeItem('tt_lic_bloq')}catch(e){}var x=document.getElementById('tt-lic-block');if(x)x.remove();var b=document.getElementById('tt-lic-warn');if(j.dias<=5){if(!b){b=document.createElement('div');b.id='tt-lic-warn';b.style.cssText='position:fixed;left:0;right:0;bottom:0;z-index:2147483000;background:#b45309;color:#fff;font:13px Arial;text-align:center;padding:6px';(document.body||document.documentElement).appendChild(b);}b.textContent='La licencia vence el '+j.vence+' (quedan '+j.dias+' dias)';}else if(b){b.remove();}}).catch(function(){});}
try{if(localStorage.getItem('tt_lic_bloq')==='1')window.addEventListener('DOMContentLoaded',bloquear)}catch(e){}
setInterval(aviso,300000);window.addEventListener('load',aviso);})();</script>
'@

function Inyectar-Guardian([string]$html) {
    $i = $html.IndexOf("<head>")
    if ($i -lt 0) { return $html }
    return $html.Insert($i + 6, $global:licGuardJs)
}

# ------------------------------------------------------------------
# Funciones de respuesta HTTP
# ------------------------------------------------------------------
function Enviar-Respuesta {
    param($Context, [string]$Body, [string]$ContentType = "text/html; charset=utf-8", [int]$StatusCode = 200)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Context.Response.StatusCode = $StatusCode
    $Context.Response.ContentType = $ContentType
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.OutputStream.Close()
}

# ------------------------------------------------------------------
# Servidor
# ------------------------------------------------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$port/")

try {
    $listener.Start()
} catch {
    Write-Host ""
    Write-Host "No se pudo iniciar el servidor en el puerto $port."
    Write-Host "Casi siempre es porque falta ejecutar como Administrador:"
    Write-Host "  - Clic derecho sobre PowerShell -> Ejecutar como administrador"
    Write-Host "  - Vuelve a correr este script desde ahi."
    Write-Host ""
    Write-Host "Detalle tecnico: $_"
    Start-Sleep -Seconds 20
    exit
}

try {
    $reglaNombre = "ServidorPedidosTotoTools"
    # Se recrea siempre: asi se corrigen reglas viejas que solo valian para red Privada
    Get-NetFirewallRule -DisplayName $reglaNombre -ErrorAction SilentlyContinue | Remove-NetFirewallRule -ErrorAction SilentlyContinue
    New-NetFirewallRule -DisplayName $reglaNombre -Direction Inbound -Protocol TCP -LocalPort $port -Action Allow -Profile Any -Enabled True | Out-Null
    Write-Host "Regla de Firewall lista para el puerto $port (todas las redes)."
    # Los hotspots y Wi-Fi nuevos llegan como "Publica"; se pasan a "Privada" para que los celulares puedan entrar
    Get-NetConnectionProfile -ErrorAction SilentlyContinue | Where-Object { $_.NetworkCategory -eq 'Public' } | Set-NetConnectionProfile -NetworkCategory Private -ErrorAction SilentlyContinue
} catch {
    Write-Host "Aviso: no se pudo crear la regla de Firewall automaticamente ($_)."
    Write-Host "Si los moviles no logran conectar, revisa el Firewall de Windows manualmente."
}

Write-Host "=================================================="
Write-Host " Servidor de Pedidos iniciado (puerto $port)"
Write-Host "=================================================="
# IP fija automatica (si la activaste en Ajustes -> Red / IP fija): se aplica en la red a la que estes conectado
if ($global:redConfig.auto -and [int]$global:redConfig.ultimoOcteto -gt 0) {
    $resIpAuto = Aplicar-IpFija ([int]$global:redConfig.ultimoOcteto)
    if ($resIpAuto.ok) { Write-Host " IP fija aplicada: $($resIpAuto.ip)" } else { Write-Host " Aviso IP fija: $($resIpAuto.error)" }
}
$patronesAdaptadorVirtual = 'Virtual|VPN|Loopback|Hyper-V|VMware|VirtualBox|Tailscale|ZeroTier|Npcap|TAP-Windows|Bluetooth'
$ips = @()
try {
    $adaptadoresUp = Get-NetAdapter -ErrorAction SilentlyContinue | Where-Object {
        $_.Status -eq 'Up' -and $_.InterfaceDescription -notmatch $patronesAdaptadorVirtual -and $_.Name -notmatch $patronesAdaptadorVirtual
    }
    foreach ($ad in $adaptadoresUp) {
        $ips += Get-NetIPAddress -InterfaceIndex $ad.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
            $_.IPAddress -ne "127.0.0.1" -and $_.IPAddress -notlike "169.254*"
        }
    }
} catch { $ips = @() }
if (-not $ips -or $ips.Count -eq 0) {
    # Respaldo: si el filtro por adaptador no encontro nada (drivers raros,
    # permisos, etc.), se cae al metodo simple anterior antes que dejar
    # $global:ipLan en "localhost".
    $ips = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object {
        $_.IPAddress -ne "127.0.0.1" -and $_.IPAddress -notlike "169.254*"
    })
}
# Si hay varias, se prefiere la de Wi-Fi (asi se conecta normalmente el
# celular, sea a un router o al hotspot de otro telefono).
# @() obligatorio: con una sola IP PowerShell 5.1 devuelve un objeto suelto sin .Count
$ips = @($ips | Sort-Object -Property @{ Expression = { if ($_.InterfaceAlias -match 'Wi-?Fi|Wireless') { 0 } else { 1 } } })
foreach ($ip in $ips) {
    Write-Host " Vendedores -> http://$($ip.IPAddress):$port/vendedor"
}
if (@($ips).Count -gt 0) {
    $global:ipLan = [string]@($ips)[0].IPAddress
} else {
    Write-Host " Aviso: no se detecto ninguna IP de red valida. El QR/enlace para clientes va a fallar hasta que conectes la PC a una red."
}
Write-Host " Panel en esta PC -> http://localhost:$port/"
Write-Host " Impresora configurada: $nombreImpresora"
if ($global:catalogoInfo.cargado) {
    Write-Host " Catalogo: $($global:catalogoInfo.cantidad) productos desde $($global:configExcel.ruta)"
} elseif (-not $global:configExcel.mapeo) {
    Write-Host " Catalogo: todavia no configurado. Se configura desde el Panel (se abre solo)."
} else {
    Write-Host " Catalogo: NO se pudo cargar ($($global:catalogoInfo.error))"
}
Write-Host "=================================================="
Write-Host "Deja esta ventana abierta. Para detener: Ctrl+C"
Write-Host ""

# Se abre como VENTANA DE APLICACION (sin pestanas ni barra de direcciones, con su propio
# boton en la barra de tareas) usando Edge o Chrome en modo "--app". Si no hay ninguno,
# se abre el navegador de siempre.
function Abrir-PanelComoApp {
    $urlPanel = "http://localhost:$port/"
    $candidatos = @(
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"),
        (Join-Path $env:ProgramFiles "Google\Chrome\Application\chrome.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Google\Chrome\Application\chrome.exe"),
        (Join-Path $env:LOCALAPPDATA "Google\Chrome\Application\chrome.exe")
    )
    foreach ($exe in $candidatos) {
        if ($exe -and (Test-Path $exe)) {
            try {
                Start-Process -FilePath $exe -ArgumentList @("--app=$urlPanel", "--window-size=1280,800", "--no-first-run")
                return $true
            } catch {}
        }
    }
    try { Start-Process $urlPanel; return $true } catch { return $false }
}
if (-not (Abrir-PanelComoApp)) {
    Write-Host "Aviso: no se pudo abrir el navegador automaticamente. Abre http://localhost:$port/ manualmente."
}

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
    } catch {
        break
    }

    # ---- LICENCIA: sin licencia vigente NO se sirve nada (ni vendedores, ni clientes, ni la PC) ----
    $request = $context.Request
    $path = $request.Url.AbsolutePath
    $method = $request.HttpMethod
    $origen = $request.RemoteEndPoint
    # ---- CORS: solo para la app APK del vendedor (su origen es http://localhost) ----
    try {
        $orig = [string]$request.Headers["Origin"]
        if ($orig -eq "http://localhost" -or $orig -eq "https://localhost" -or $orig -eq "capacitor://localhost") {
            $context.Response.Headers["Access-Control-Allow-Origin"] = $orig
            $context.Response.Headers["Vary"] = "Origin"
            $context.Response.Headers["Access-Control-Allow-Headers"] = "Content-Type, X-Vendedor, X-Pos-Token"
            $context.Response.Headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
            $context.Response.Headers["Access-Control-Expose-Headers"] = "X-Nombre-Archivo"
            $context.Response.Headers["Access-Control-Allow-Private-Network"] = "true"
        }
    } catch {}
    if ($method -eq "OPTIONS") {
        try { $context.Response.StatusCode = 204; $context.Response.ContentLength64 = 0; $context.Response.OutputStream.Close() } catch {}
        continue
    }
    if (-not [TotoLic.Gate]::IsOpen()) {
        try {
            if ($method -ne "GET" -or $path.StartsWith("/api/")) {
                Enviar-Respuesta -Context $context -Body '{"ok":false,"licencia":false,"error":"Licencia vencida o no valida. Contacta al proveedor."}' -ContentType "application/json; charset=utf-8" -StatusCode 402
            } else {
                Enviar-Respuesta -Context $context -Body $global:htmlSinLicencia -StatusCode 402
            }
        } catch {}
        continue
    }
    if ($method -eq "GET" -and $path -eq "/api/licencia") {
        try { Enviar-Respuesta -Context $context -Body (@{ ok = $true; dias = [int][TotoLic.Gate]::Dias(); vence = [TotoLic.Gate]::VenceTexto() } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" } catch {}
        continue
    }

    Revisar-CambioCatalogo

    # Las consultas de sincronizacion del Gestor de Almacenes (cada ~10 s por telefono) no se anotan en pantalla para no llenar la consola.
    if (-not ($method -eq "GET" -and ($path.StartsWith("/api/almacen/") -or $path -eq "/api/fotos/skus"))) {
        Write-Host "[$((Get-Date).ToString('HH:mm:ss'))] $method $path  <- $origen"
    }

    try {
        $permReq = Permiso-De-Ruta $method $path
        $vendHdr = ""
        try { $vendHdr = [uri]::UnescapeDataString([string]$request.Headers["X-Vendedor"]) } catch { $vendHdr = "" }
        $vendHdr = $vendHdr.Trim()
        $esLocalGate = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
        if ($permReq -and $vendHdr -and (-not $esLocalGate) -and (-not (Tiene-Permiso $vendHdr $permReq))) {
            Enviar-Json $context @{ ok = $false; error = "Tu usuario no tiene permiso para hacer esto. Pidele al encargado que lo active desde la PC."; sinPermiso = $true; permiso = $permReq } 403

        } elseif ($method -eq "GET" -and ($path -eq "/" -or $path -eq "/pc")) {
            Enviar-Respuesta -Context $context -Body (Inyectar-Guardian $htmlPC)

        } elseif ($method -eq "GET" -and $path -eq "/vendedor") {
            Enviar-Respuesta -Context $context -Body (Inyectar-Guardian $htmlVendedor)

        } elseif ($method -eq "GET" -and $path -eq "/metricas") {
            Enviar-Respuesta -Context $context -Body (Inyectar-Guardian $htmlMetricas)

        } elseif ($method -eq "GET" -and $path -eq "/api/metricas") {
            $diasQ = 1
            if ($request.QueryString["dias"]) { try { $diasQ = [int]$request.QueryString["dias"] } catch { $diasQ = 1 } }
            if ($diasQ -lt 1) { $diasQ = 1 }
            if ($diasQ -gt 90) { $diasQ = 90 }
            try {
                $m = Calcular-Metricas $diasQ
                Enviar-Respuesta -Context $context -Body ($m | ConvertTo-Json -Depth 6) -ContentType "application/json; charset=utf-8"
            } catch {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "$($_.Exception.Message)" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
            }

        } elseif ($method -eq "GET" -and $path -eq "/xlsx.js") {
            if ($global:xlsxLibJs) {
                Enviar-Respuesta -Context $context -Body $global:xlsxLibJs -ContentType "application/javascript; charset=utf-8"
            } else {
                Enviar-Respuesta -Context $context -Body "// falta xlsx_full_min.js junto al script" -ContentType "application/javascript; charset=utf-8" -StatusCode 404
            }

        } elseif ($method -eq "GET" -and $path -eq "/manifest-vendedor.json") {
            Enviar-Respuesta -Context $context -Body $global:manifestVendedor -ContentType "application/manifest+json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/manifest-pc.json") {
            Enviar-Respuesta -Context $context -Body $global:manifestPC -ContentType "application/manifest+json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/sw.js") {
            Enviar-Respuesta -Context $context -Body $global:swJs -ContentType "application/javascript; charset=utf-8"

        } elseif ($method -eq "GET" -and ($path -eq "/icon-192.png" -or $path -eq "/icon-512.png" -or $path -eq "/icon-pc-192.png" -or $path -eq "/icon-pc-512.png")) {
            # /icon-* = carrito (app del movil); /icon-pc-* = martillo (app de la PC)
            $bytesIcono = switch ($path) {
                "/icon-192.png"    { $global:icon192Bytes }
                "/icon-512.png"    { $global:icon512Bytes }
                "/icon-pc-192.png" { $global:iconPc192Bytes }
                default            { $global:iconPc512Bytes }
            }
            $context.Response.StatusCode = 200
            $context.Response.ContentType = "image/png"
            $context.Response.ContentLength64 = $bytesIcono.Length
            $context.Response.OutputStream.Write($bytesIcono, 0, $bytesIcono.Length)
            $context.Response.OutputStream.Close()

        } elseif ($method -eq "GET" -and $path -eq "/api/ip") {
            Enviar-Respuesta -Context $context -Body (@{ ip = $global:ipLan; puerto = $port } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/config") {
            Enviar-Respuesta -Context $context -Body ($global:configApp | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/config") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json
            if ($data.PSObject.Properties.Name -contains 'ocultarSinStock') {
                $global:configApp.ocultarSinStock = [bool]$data.ocultarSinStock
            }
            if ($data.PSObject.Properties.Name -contains 'tasaDolar') {
                $tasaNueva = 0.0
                try { $tasaNueva = [double]$data.tasaDolar } catch { $tasaNueva = 0.0 }
                if ($tasaNueva -lt 0) { $tasaNueva = 0.0 }
                $global:configApp.tasaDolar = $tasaNueva
            }
            if ($data.PSObject.Properties.Name -contains 'permitirDescuentos') {
                $global:configApp.permitirDescuentos = [bool]$data.permitirDescuentos
            }
            if ($data.PSObject.Properties.Name -contains 'umbralStockBajo') {
                $umbralNuevo = 3
                try { $umbralNuevo = [int]$data.umbralStockBajo } catch { $umbralNuevo = 3 }
                if ($umbralNuevo -lt 0) { $umbralNuevo = 0 }
                $global:configApp.umbralStockBajo = $umbralNuevo
            }
            if ($data.PSObject.Properties.Name -contains 'autoservicioDestino') {
                $global:configApp.autoservicioDestino = if ($data.autoservicioDestino -eq 'vendedor') { "vendedor" } else { "pc" }
            }
            if ($data.PSObject.Properties.Name -contains 'wifiSSID') {
                $global:configApp.wifiSSID = "$($data.wifiSSID)"
            }
            if ($data.PSObject.Properties.Name -contains 'wifiClave') {
                $global:configApp.wifiClave = "$($data.wifiClave)"
            }
            Guardar-ConfigApp
            Enviar-Respuesta -Context $context -Body ($global:configApp | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/stockbajo") {
            # Productos con existencia por debajo (o igual) del umbral configurado
            # en Ajustes. umbral 0 desactiva la alerta (no se avisa de nada).
            $umbral = [int]$global:configApp.umbralStockBajo
            $bajos = @()
            if ($true) {
                $bajos = @($global:catalogo | Where-Object { $mm = Minimo-Efectivo $_; $_.stock -ne $null -and $mm -gt 0 -and [double]$_.stock -le $mm } |
                    Sort-Object -Property @{ Expression = { [double]$_.stock } } |
                    ForEach-Object { [pscustomobject]@{ sku = $_.sku; nombre = $_.nombre; stock = $_.stock } })
            }
            Enviar-Respuesta -Context $context -Body (@{ umbral = $umbral; productos = $bajos } | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/almacen/ping") {
            $ci = $global:catalogoInfo
            Enviar-Json $context @{ ok = $true; version = 1; ip = $global:ipLan; puerto = $port; catalogo = @{ cargado = [bool]$ci.cargado; cantidad = [int]$ci.cantidad; ultimaCarga = $ci.ultimaCarga; error = $ci.error } }

        } elseif ($method -eq "GET" -and $path -eq "/api/almacen/mov") {
            Enviar-Respuesta -Context $context -Body (Almacen-ListaMovJson) -ContentType "application/json; charset=utf-8"

        } elseif ($path -eq "/api/almacen/mov/archivo" -and ($method -eq "GET" -or $method -eq "POST")) {
            $nomMov = Almacen-NombreValido ([string]$request.QueryString["name"])
            if (-not $nomMov) {
                Enviar-Json $context @{ ok = $false; error = "Nombre de archivo no valido." } 400
            } elseif ($method -eq "GET") {
                $txtMov = Almacen-Leer (Join-Path $almacenSyncMovDir $nomMov)
                if ($null -eq $txtMov) { Enviar-Json $context @{ ok = $false; error = "No existe." } 404 }
                else { Enviar-Respuesta -Context $context -Body $txtMov -ContentType "application/json; charset=utf-8" }
            } else {
                $cuerpoMov = ([string](Leer-CuerpoTexto $request)).Trim()
                if ($cuerpoMov.Length -lt 2 -or -not $cuerpoMov.StartsWith("{") -or $cuerpoMov.Length -gt 30000000) {
                    Enviar-Json $context @{ ok = $false; error = "Contenido no valido." } 400
                } else {
                    Almacen-Escribir (Join-Path $almacenSyncMovDir $nomMov) $cuerpoMov
                    Enviar-Json $context @{ ok = $true; sha = (Almacen-Sha $cuerpoMov) }
                }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/almacen/respaldo") {
            $txtResp = Almacen-Leer $almacenRespaldoPath
            $cuerpoResp = if ($null -eq $txtResp) { '{"sha":null,"contenido":null}' } else { '{"sha":"' + (Almacen-Sha $txtResp) + '","contenido":' + $txtResp + '}' }
            Enviar-Respuesta -Context $context -Body $cuerpoResp -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/almacen/respaldo") {
            $cuerpoResp = ([string](Leer-CuerpoTexto $request)).Trim()
            if ($cuerpoResp.Length -lt 2 -or -not $cuerpoResp.StartsWith("{") -or $cuerpoResp.Length -gt 30000000) {
                Enviar-Json $context @{ ok = $false; error = "Contenido no valido." } 400
            } else {
                $actualResp = Almacen-Leer $almacenRespaldoPath
                $shaActualResp = if ($null -ne $actualResp) { Almacen-Sha $actualResp } else { "" }
                $baseResp = [string]$request.QueryString["base"]
                # Control de concurrencia: si otro telefono guardo despues de que este leyo, se rechaza y el telefono reintenta.
                if ($baseResp -ne $shaActualResp) {
                    Enviar-Json $context @{ ok = $false; conflicto = $true; sha = $shaActualResp } 409
                } else {
                    Almacen-Escribir $almacenRespaldoPath $cuerpoResp
                    Enviar-Json $context @{ ok = $true; sha = (Almacen-Sha $cuerpoResp) }
                }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/catalogo") {
            $json = if ($global:catalogo.Count -eq 0) { "[]" } else {
                $j = $global:catalogo | ConvertTo-Json -Depth 6
                if ($global:catalogo.Count -eq 1) { "[" + $j + "]" } else { $j }
            }
            Enviar-Respuesta -Context $context -Body $json -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/catalogo/estado") {
            Enviar-Respuesta -Context $context -Body ($global:catalogoInfo | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/catalogo/recargar") {
            $global:catalogoMTimeProcesada = $null
            Revisar-CambioCatalogo
            Enviar-Respuesta -Context $context -Body ($global:catalogoInfo | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/catalogo/config") {
            Enviar-Respuesta -Context $context -Body ($global:configExcel | ConvertTo-Json -Depth 10) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/catalogo/config") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json

            if ($data.PSObject.Properties.Name -contains 'ruta' -and $data.ruta) {
                $global:configExcel.ruta = [string]$data.ruta
                $global:catalogoInfo.archivo = $global:configExcel.ruta
                $global:catalogoMTimeProcesada = $null
                $global:catalogoInfo.cargado = $false
                $global:catalogoInfo.error = $null
            }
            if ($data.PSObject.Properties.Name -contains 'mapeo' -and $data.mapeo) {
                $global:configExcel.mapeo = $data.mapeo
                $global:catalogoInfo.mapeoConfigurado = $true
                $global:catalogoMTimeProcesada = $null
            }
            Guardar-ConfigExcel
            Revisar-CambioCatalogo
            Enviar-Respuesta -Context $context -Body ($global:catalogoInfo | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/catalogo/archivo") {
            $ruta = $global:configExcel.ruta
            if (-not $ruta -or -not (Test-Path $ruta)) {
                Enviar-Respuesta -Context $context -Body "No se encontro el archivo Excel configurado." -ContentType "text/plain; charset=utf-8" -StatusCode 404
            } else {
                try {
                    $bytesArchivo = [System.IO.File]::ReadAllBytes($ruta)
                    $nombreArchivo = [System.IO.Path]::GetFileName($ruta)
                    $context.Response.StatusCode = 200
                    $context.Response.ContentType = "application/octet-stream"
                    $context.Response.Headers.Add("Content-Disposition", "attachment; filename=`"$nombreArchivo`"")
                    $context.Response.Headers.Add("X-Nombre-Archivo", $nombreArchivo)
                    $context.Response.ContentLength64 = $bytesArchivo.Length
                    $context.Response.OutputStream.Write($bytesArchivo, 0, $bytesArchivo.Length)
                    $context.Response.OutputStream.Close()
                } catch {
                    Enviar-Respuesta -Context $context -Body "Error leyendo el archivo: $_" -ContentType "text/plain; charset=utf-8" -StatusCode 500
                }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/catalogo/importar") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            try {
                $data = $bodyText | ConvertFrom-Json
                $productos = @($data.productos)

                # Indice del catalogo ANTERIOR por SKU, para no perder ventas
                # ya descontadas cuando el Excel cambia por otro motivo.
                $anteriorPorSku = @{}
                foreach ($pAnt in $global:catalogo) {
                    if ($pAnt.sku) { $anteriorPorSku[[string]$pAnt.sku] = $pAnt }
                }

                $nuevo = New-Object System.Collections.ArrayList
                foreach ($p in $productos) {
                    $nombre = ([string]$p.nombre).Trim()
                    if ([string]::IsNullOrWhiteSpace($nombre)) { continue }
                    $sku = if ($p.sku) { ([string]$p.sku).Trim() } else { "" }
                    $precio = 0.0
                    if ($p.precio -ne $null) { [double]::TryParse([string]$p.precio, [ref]$precio) | Out-Null }

                    $stockExcel = $null
                    if ($p.stock -ne $null -and $p.stock -ne "") {
                        $stockVal = 0.0
                        if ([double]::TryParse([string]$p.stock, [ref]$stockVal)) { $stockExcel = $stockVal }
                    }

                    # Por defecto: el Excel es la nueva verdad (producto nuevo,
                    # o sin control de stock).
                    $stockBase = $stockExcel
                    $vendido   = 0.0
                    $stockEfectivo = $stockExcel

                    $anterior = if ($sku) { $anteriorPorSku[$sku] } else { $null }
                    if ($anterior -and $stockExcel -ne $null -and $anterior.stockBase -ne $null) {
                        if ([math]::Abs([double]$anterior.stockBase - [double]$stockExcel) -lt 0.0001) {
                            # La cantidad en el Excel NO cambio -> se preservan
                            # las ventas hechas desde el movil desde la ultima
                            # vez que si cambio.
                            $stockBase = [double]$anterior.stockBase
                            $vendido = [double]$anterior.vendido
                            $stockEfectivo = $stockBase - $vendido
                            if ($stockEfectivo -lt 0) { $stockEfectivo = 0 }
                        }
                        # Si SI cambio (reabasteciste / corregiste el conteo),
                        # esa cifra del Excel se toma como el nuevo punto de
                        # partida y se reinicia lo "vendido pendiente".
                    }

                    [void]$nuevo.Add([pscustomobject]@{
                        sku = $sku
                        nombre = $nombre
                        precio = [math]::Round($precio,2)
                        stock = $stockEfectivo
                        stockBase = $stockBase
                        vendido = $vendido
                    })
                }

                Avisar-CambiosCatalogo $anteriorPorSku $nuevo
                Registrar-ProductosNuevos $nuevo
                $global:catalogo = $nuevo
                $global:catalogoInfo.cargado = $true
                $global:catalogoInfo.ultimaCarga = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $global:catalogoInfo.error = $null
                $global:catalogoInfo.cantidad = $nuevo.Count
                $global:catalogoInfo.necesitaRecarga = $false
                if (Test-Path $global:configExcel.ruta) {
                    $global:catalogoMTimeProcesada = (Get-Item $global:configExcel.ruta).LastWriteTimeUtc
                }
                Write-Host "Catalogo actualizado desde el Panel: $($nuevo.Count) productos"
                Enviar-Respuesta -Context $context -Body ($global:catalogoInfo | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            } catch {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "$_" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/fotos/skus") {
            # Lista de SKU que SI tienen foto: los paneles solo piden /foto/<sku>.jpg de esos,
            # asi un producto sin imagen no provoca peticiones fallidas ni parpadeos.
            $listaSkusFoto = New-Object System.Collections.ArrayList
            $carpetaFotosLista = Join-Path $scriptDir "fotos_catalogo"
            if (Test-Path $carpetaFotosLista) {
                foreach ($ffLista in @(Get-ChildItem -Path $carpetaFotosLista -Filter "*.jpg" -File -ErrorAction SilentlyContinue)) {
                    [void]$listaSkusFoto.Add([System.IO.Path]::GetFileNameWithoutExtension($ffLista.Name))
                }
            }
            Enviar-Respuesta -Context $context -Body (ConvertTo-Json -InputObject @($listaSkusFoto) -Compress) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -match "^/foto/grande/([^/]+)\.jpg$") {
            $skuFoto = [System.Uri]::UnescapeDataString($matches[1])
            $rutaFoto = Join-Path (Join-Path (Join-Path $scriptDir "fotos_catalogo") "grande") ("$skuFoto.jpg")
            if (-not (Test-Path $rutaFoto)) {
                # Fotos importadas antes de tener version "grande": se manda la chica.
                $rutaFoto = Join-Path (Join-Path $scriptDir "fotos_catalogo") ("$skuFoto.jpg")
            }
            if (Test-Path $rutaFoto) {
                try {
                    $bytesFoto = [System.IO.File]::ReadAllBytes($rutaFoto)
                    $context.Response.StatusCode = 200
                    $context.Response.ContentType = "image/jpeg"
                    $context.Response.Headers.Add("Cache-Control", "public, max-age=86400")
                    $context.Response.ContentLength64 = $bytesFoto.Length
                    $context.Response.OutputStream.Write($bytesFoto, 0, $bytesFoto.Length)
                    $context.Response.OutputStream.Close()
                } catch {
                    $context.Response.StatusCode = 404
                    $context.Response.OutputStream.Close()
                }
            } else {
                $context.Response.StatusCode = 404
                $context.Response.OutputStream.Close()
            }

        } elseif ($method -eq "GET" -and $path -match "^/foto/([^/]+)\.jpg$") {
            $skuFoto = [System.Uri]::UnescapeDataString($matches[1])
            $rutaFoto = Join-Path (Join-Path $scriptDir "fotos_catalogo") ("$skuFoto.jpg")
            if (Test-Path $rutaFoto) {
                try {
                    $bytesFoto = [System.IO.File]::ReadAllBytes($rutaFoto)
                    $context.Response.StatusCode = 200
                    $context.Response.ContentType = "image/jpeg"
                    $context.Response.Headers.Add("Cache-Control", "public, max-age=86400")
                    $context.Response.ContentLength64 = $bytesFoto.Length
                    $context.Response.OutputStream.Write($bytesFoto, 0, $bytesFoto.Length)
                    $context.Response.OutputStream.Close()
                } catch {
                    $context.Response.StatusCode = 404
                    $context.Response.OutputStream.Close()
                }
            } else {
                $context.Response.StatusCode = 404
                $context.Response.OutputStream.Close()
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/fotos/importar") {
            $resultado = Importar-FotosCatalogo
            $codigo = if ($resultado.ok) { 200 } else { 400 }
            Enviar-Respuesta -Context $context -Body ($resultado | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode $codigo

        } elseif ($method -eq "POST" -and $path -eq "/api/fotos/subir") {
            # Recibe el .zip de "Copia de seguridad" subido directo desde el
            # movil (sin pasar por la PC), lo guarda en esta carpeta con
            # nombre reconocible y reusa la misma logica de emparejado.
            try {
                $ms = New-Object System.IO.MemoryStream
                $request.InputStream.CopyTo($ms)
                $bytesZip = $ms.ToArray()
                $ms.Close()
                if ($bytesZip.Length -eq 0) { throw "Archivo vacio." }
                $nombreZip = "backup-catalogo-movil-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".zip"
                $rutaZip = Join-Path $scriptDir $nombreZip
                [System.IO.File]::WriteAllBytes($rutaZip, $bytesZip)
                $resultado = Importar-FotosCatalogo
                $codigo = if ($resultado.ok) { 200 } else { 400 }
                Enviar-Respuesta -Context $context -Body ($resultado | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode $codigo
            } catch {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "$_" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/fotos/estado") {
            $carpetaFotos = Join-Path $scriptDir "fotos_catalogo"
            $cantidadFotos = 0
            if (Test-Path $carpetaFotos) {
                $cantidadFotos = @(Get-ChildItem -Path $carpetaFotos -Filter "*.jpg" -File -ErrorAction SilentlyContinue).Count
            }
            Enviar-Respuesta -Context $context -Body (@{ cantidadFotos = $cantidadFotos } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/dia/cerrar") {
            $ok = Cerrar-Dia
            if ($ok) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $true; restantes = $global:pedidos.Count } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            } else {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "No se pudo archivar el historial." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/ping") {
            # El movil llama esto cada 10s para saber si la PC responde (indicador de red).
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; hora = (Get-Date).ToString("HH:mm:ss"); ts = [int64][DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds() } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/vendedores/ping") {
            # El movil del vendedor llama esto cada pocos segundos mientras tenga
            # su nombre puesto, para que la PC sepa quien esta conectado ahora.
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }
            $nombreVend = if ($data -and $data.nombre) { ([string]$data.nombre).Trim() } else { "" }
            if ($nombreVend) { $global:vendedoresConectados[$nombreVend] = Get-Date }
            Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/vendedores") {
            # La PC pide la lista de vendedores conectados ahora mismo, para el
            # selector de "Nuevo pedido para vendedor".
            $limite = (Get-Date).AddSeconds(-$global:segundosVendedorConectado)
            $conectados = @($global:vendedoresConectados.Keys | Where-Object { $global:vendedoresConectados[$_] -ge $limite } | Sort-Object)
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; vendedores = $conectados } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/pines") {
            # Lista de vendedores registrados (nunca se manda el PIN, solo el
            # nombre tal como lo escribio la PC). La usan tanto el Panel
            # ("PIN de vendedores") como el login del movil (para elegir de
            # la lista en vez de escribir el nombre).
            $conPin = @($global:pinesVendedores.Keys | ForEach-Object { $global:pinesVendedores[$_].nombre } | Sort-Object)
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; vendedoresConPin = $conPin } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/pines/verificar") {
            # El movil pregunta si este vendedor tiene PIN y si el que escribio es
            # correcto (para salir del modo cliente). Nunca se devuelve el PIN.
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json
            $vendV = if ($data -and $data.vendedor) { ([string]$data.vendedor).Trim() } else { "" }
            $claveV = $vendV.ToLowerInvariant()
            $existeV = ($vendV -ne "") -and $global:pinesVendedores.ContainsKey($claveV)
            $tienePinV = $existeV -and (-not [string]::IsNullOrEmpty($global:pinesVendedores[$claveV].pin))
            $largoPinV = if ($tienePinV) { ([string]$global:pinesVendedores[$claveV].pin).Length } else { 0 }
            $validoV = Pin-Valido $vendV $data.pin
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; existe = $existeV; tienePin = $tienePinV; largoPin = $largoPinV; valido = $validoV } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/pines") {
            # Crear un vendedor nuevo SOLO se puede desde la PC (loopback), con
            # nombre que no este en uso y PIN obligatorio (4 a 6 numeros) -- asi
            # nadie puede inventarse un nombre de vendedor desde el movil.
            # Un vendedor que YA existe puede cambiarse su propio PIN desde el
            # movil (Ajustes), pero hace falta el PIN actual correcto; el PIN
            # sigue siendo obligatorio, no se puede dejar vacio desde el movil.
            # Desde la PC, mandar el PIN vacio para un vendedor que YA existe lo
            # ELIMINA (ya no tiene sentido un vendedor sin PIN).
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json
            $vend = if ($data -and $data.vendedor) { ([string]$data.vendedor).Trim() } else { "" }
            $pinNuevo = if ($data -and $data.pin) { [string]$data.pin } else { "" }
            $esLocalPin = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $clave = $vend.ToLowerInvariant()
            $yaExiste = $vend -and $global:pinesVendedores.ContainsKey($clave)
            if (-not $vend) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Falta el nombre del vendedor." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ((-not $yaExiste) -and (-not $esLocalPin)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Ese vendedor no existe. Pidele al encargado que te cree el usuario desde la PC."; requiereCreacionPC = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
            } elseif ($pinNuevo -and $pinNuevo -notmatch "^\d{4,6}$") {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "El PIN debe tener de 4 a 6 numeros." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ((-not $yaExiste) -and (-not $pinNuevo)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "El PIN es obligatorio para crear un vendedor nuevo." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($yaExiste -and (-not $esLocalPin) -and (-not $pinNuevo)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "El PIN es obligatorio, no se puede dejar vacio." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($yaExiste -and (-not $esLocalPin) -and -not (Pin-Valido $vend $data.pinActual)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "PIN actual incorrecto."; requierePin = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
            } elseif ($yaExiste -and $esLocalPin -and (-not $pinNuevo)) {
                # La PC manda el PIN vacio para un vendedor existente: se elimina.
                $global:pinesVendedores.Remove($clave) | Out-Null
                Guardar-Pines
                Enviar-Respuesta -Context $context -Body (@{ ok = $true; eliminado = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            } else {
                $global:pinesVendedores[$clave] = @{ nombre = $vend; pin = $pinNuevo }
                Guardar-Pines
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/pedidos/asignar") {
            # La PC arma un pedido (productos y cantidades) y lo manda a un
            # vendedor concreto; su movil lo recoge solo en el proximo aviso.
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json
            $vendedorDestino = if ($data -and $data.vendedor) { ([string]$data.vendedor).Trim() } else { "" }
            $paraTodos = [bool]($data -and $data.todos)
            $itemsAsignados = @($data.items)
            $clienteOrigen = if ($data -and $data.cliente) { ([string]$data.cliente).Trim() } else { "" }
            # La nota solo la puede poner la propia PC (un cliente no puede mandar mensajes a los vendedores).
            $notaAsig = ""
            if ($data -and $data.nota -and [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)) { $notaAsig = Limpiar-TextoMensaje $data.nota }
            $claveEnvioAsig = if ($data -and $data.claveEnvio) { "asig|" + [string]$data.claveEnvio } else { "" }
            $minVigAsig = 60
            try { $mvA = [int]$data.minutosVigencia; if ($mvA -ge 5 -and $mvA -le 1440) { $minVigAsig = $mvA } } catch {}
            $vigenciaAsig = (Get-Date).AddMinutes($minVigAsig).ToString("yyyy-MM-dd HH:mm:ss")
            if ((-not $vendedorDestino) -and (-not $paraTodos)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Falta elegir el vendedor." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($itemsAsignados.Count -eq 0) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Agrega al menos un producto." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($respPreviaAsig = (Clave-EnvioBuscar $claveEnvioAsig)) {
                # Mismo envio repetido: se devuelve el pedido que ya se creo, sin crear otro.
                Enviar-Respuesta -Context $context -Body $respPreviaAsig -ContentType "application/json; charset=utf-8"
            } else {
                $asignado = [ordered]@{
                    id            = $global:nextIdAsignado
                    vendedor      = $vendedorDestino
                    cliente       = $clienteOrigen
                    todos         = $paraTodos
                    entregadoA    = New-Object System.Collections.ArrayList
                    tomadoPor     = $null
                    horaTomado    = $null
                    items         = $itemsAsignados
                    hora          = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    entregado     = $false
                    visto         = $false
                    horaVisto     = $null
                    retirado      = $false
                    horaRetirado  = $null
                    vence         = $vigenciaAsig
                    nota          = $notaAsig
                }
                $global:nextIdAsignado++
                [void]$global:pedidosAsignados.Add([pscustomobject]$asignado)
                Guardar-Asignados
                $respAsigJson = (@{ ok = $true; id = $asignado.id; repetido = $true } | ConvertTo-Json -Compress)
                Clave-EnvioGuardar $claveEnvioAsig $respAsigJson
                Enviar-Respuesta -Context $context -Body (@{ ok = $true; id = $asignado.id } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/pedidos/asignados") {
            # El movil del vendedor pregunta si la PC le armo algun pedido nuevo.
            # A cada vendedor se le entrega una sola vez (se anota su nombre en
            # entregadoA), pero NO se borra: se queda en la cola para que la PC
            # pueda ver despues si ya lo "vio" o, si es para todos, quien lo tomo.
            # Los "para todos" se entregan al mismo tiempo a todos los
            # vendedores conectados y siguen visibles para todos hasta que
            # alguno lo toma (boton "Yo lo tomo" en su movil).
            $nombreVend = if ($request.QueryString["vendedor"]) { $request.QueryString["vendedor"].Trim() } else { "" }
            $ahoraEntrega = Get-Date
            $paraEste = @($global:pedidosAsignados | Where-Object {
                (-not $_.tomadoPor) -and (-not $_.retirado) -and ((-not $_.vence) -or ([datetime]$_.vence -gt $ahoraEntrega)) -and ($_.entregadoA -notcontains $nombreVend) -and
                (($_.todos) -or ($_.vendedor -eq $nombreVend))
            })
            foreach ($a in $paraEste) {
                [void]$a.entregadoA.Add($nombreVend)
                if (-not $a.todos) { $a.entregado = $true }
            }
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; pedidos = $paraEste } | ConvertTo-Json -Depth 10) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/asignados/(\d+)/visto$") {
            # El vendedor marca en su movil que ya vio un pedido que le armo la
            # caja (aunque todavia no lo haya agregado a su pedido actual).
            $idAsig = [int]$Matches[1]
            $asig = $global:pedidosAsignados | Where-Object { [int]$_.id -eq $idAsig }
            if (-not $asig) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "No encontrado." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } else {
                $asig.visto = $true
                $asig.horaVisto = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Guardar-Asignados
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/asignados/(\d+)/tomar$") {
            # Un vendedor toca "Yo lo tomo" en un pedido "para todos". Solo el
            # primero que llegue se lo queda; a partir de ahi desaparece para
            # los demas y la PC ve quien lo va a recibir.
            $idAsig = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = if ($bodyText -and $bodyText.Trim().Length -gt 0) { $bodyText | ConvertFrom-Json } else { $null }
            $nombreVend = if ($data -and $data.vendedor) { ([string]$data.vendedor).Trim() } else { "" }
            $asig = $global:pedidosAsignados | Where-Object { [int]$_.id -eq $idAsig }
            if (-not $asig) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "No encontrado." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif (-not $nombreVend) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Falta el nombre del vendedor." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($asig.tomadoPor) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Ya lo tomo " + $asig.tomadoPor + "."; tomadoPor = $asig.tomadoPor } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } else {
                $asig.tomadoPor  = $nombreVend
                $asig.horaTomado = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $asig.visto = $true
                $asig.horaVisto = $asig.horaTomado
                Guardar-Asignados
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/asignados/(\d+)/retirar$") {
            # La PC "deshace" un pedido que armo por error (para todos o para
            # un vendedor concreto), siempre que TODAVIA nadie lo haya tomado.
            # No se borra de una vez: se marca "retirado" para que, si algun
            # movil ya lo tenia en pantalla, lo pueda quitar solo en su
            # proxima consulta (ver /api/pedidos/asignados/estado).
            $idAsig = [int]$Matches[1]
            $asig = $global:pedidosAsignados | Where-Object { [int]$_.id -eq $idAsig }
            if (-not $asig) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "No encontrado." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif ($asig.tomadoPor) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Ya lo tomo " + $asig.tomadoPor + ", ya no se puede retirar."; tomadoPor = $asig.tomadoPor } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif ($asig.retirado) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            } else {
                $asig.retirado = $true
                $asig.horaRetirado = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Guardar-Asignados
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/pedidos/asignados/estado") {
            # La PC consulta si los pedidos que ha mandado ya fueron vistos (o,
            # si eran para todos, quien se los tomo), para avisar con un
            # banner. Tambien se usa para que el movil detecte si un pedido
            # que le llego ya no vale (lo tomo otro, o la PC lo retiro).
            # Se limpian los ya vistos hace mas de 2 horas, y los retirados
            # hace mas de 2 minutos (les damos un rato para que el movil que
            # los tenia en pantalla note el cambio), para no acumular memoria.
            $ahora = Get-Date
            $limiteViejo = $ahora.AddHours(-2)
            $limiteRetirado = $ahora.AddMinutes(-2)
            $aQuitar = @($global:pedidosAsignados | Where-Object {
                ($_.visto -and $_.horaVisto -and ([datetime]$_.horaVisto) -lt $limiteViejo) -or
                ($_.retirado -and $_.horaRetirado -and ([datetime]$_.horaRetirado) -lt $limiteRetirado) -or
                ((-not $_.tomadoPor) -and $_.vence -and ([datetime]$_.vence) -lt $limiteRetirado)
            })
            if ($aQuitar.Count -gt 0) {
                foreach ($a in $aQuitar) { [void]$global:pedidosAsignados.Remove($a) }
                Guardar-Asignados
            }
            $todos = @($global:pedidosAsignados | ForEach-Object {
                $vencido = ($_.todos -and (-not $_.tomadoPor) -and (-not $_.retirado) -and (($ahora - [datetime]$_.hora).TotalMinutes -ge $global:minutosAvisoSinTomar))
                [pscustomobject]@{ id = $_.id; vendedor = $_.vendedor; todos = [bool]$_.todos; tomadoPor = $_.tomadoPor; hora = $_.hora; entregado = [bool]$_.entregado; visto = [bool]$_.visto; retirado = [bool]$_.retirado; vencido = [bool]$vencido }
            })
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; asignados = $todos } | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "GET" -and $path -eq "/api/pedidos") {
            if ($global:pedidos.Count -eq 0) {
                $json = "[]"
            } else {
                $json = $global:pedidos | ConvertTo-Json -Depth 10
                if ($global:pedidos.Count -eq 1) { $json = "[" + $json + "]" }
            }
            Enviar-Respuesta -Context $context -Body $json -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -eq "/api/pedidos") {
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $bodyText | ConvertFrom-Json

            # ---- Pedido de un CLIENTE (autoservicio): no lleva PIN ni usuario. Para que
            # no se pueda abusar, el servidor lo fuerza a "pendiente", con los precios
            # reales del catalogo y a nombre de "Cliente: <nombre>". ----
            $esAutoCli = $false
            try { $esAutoCli = ($data.autoservicio -eq $true) } catch { $esAutoCli = $false }
            $errAutoCli = ""
            if ($esAutoCli) {
                $nomCli = ([string]$data.vendedor).Trim()
                if (-not $nomCli) { $nomCli = "Cliente" }
                if ($nomCli.Length -gt 30) { $nomCli = $nomCli.Substring(0, 30) }
                $itemsCli = New-Object System.Collections.ArrayList
                foreach ($itC in @($data.items)) {
                    $skC = [string]$itC.sku
                    if ([string]::IsNullOrWhiteSpace($skC)) { continue }
                    $pC = $global:catalogo | Where-Object { $_.sku -eq $skC } | Select-Object -First 1
                    $cantC = 0.0; try { $cantC = [double]$itC.cantidad } catch { $cantC = 0.0 }
                    if (-not $pC -or $cantC -le 0) { continue }
                    $itC | Add-Member -NotePropertyName precio -NotePropertyValue ([double]$pC.precio) -Force
                    $itC | Add-Member -NotePropertyName cantidad -NotePropertyValue $cantC -Force
                    [void]$itemsCli.Add($itC)
                }
                if ($itemsCli.Count -eq 0) { $errAutoCli = "Agrega al menos un producto." }
                $forzar = @{ vendedor = ("Cliente: " + $nomCli); estado = "pendiente"; metodoPago = "Efectivo"; montoRecibido = $null; cambio = $null; pagoEfectivo = $null; pagoTransferencia = $null; origenAsignados = $null; items = @($itemsCli) }
                foreach ($kF in $forzar.Keys) { $data | Add-Member -NotePropertyName $kF -NotePropertyValue $forzar[$kF] -Force }
            }

            # Bandera en vez de "return": este bucle no es una funcion, un return
            # aqui pararia todo el servidor en vez de solo esta peticion.
            $continuarPedido = $true
            if ($errAutoCli) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = $errAutoCli } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
                $continuarPedido = $false
            }

            # Permisos del vendedor (los define la PC). Desde la propia PC no se limita nada.
            $esLocalCrear = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if ($continuarPedido -and (-not $esAutoCli) -and (-not $esLocalCrear) -and $data.vendedor) {
                $vendCrear = [string]$data.vendedor
                $noPermitido = ""
                if (-not (Tiene-Permiso $vendCrear 'crearPedidos')) {
                    $noPermitido = "Tu usuario no puede crear pedidos."
                } elseif (([string]$data.estado -eq "cobrado") -and (-not (Tiene-Permiso $vendCrear 'cobrar'))) {
                    $noPermitido = "Tu usuario no puede cobrar pedidos: mandalo como pendiente de pago."
                } elseif ((-not (Tiene-Permiso $vendCrear 'descuentos')) -and (-not $data.origenAsignados)) {
                    foreach ($itD in @($data.items)) {
                        if ([string]::IsNullOrWhiteSpace([string]$itD.sku)) { continue }
                        $prodD2 = $global:catalogo | Where-Object { $_.sku -eq [string]$itD.sku } | Select-Object -First 1
                        if ($prodD2 -and [double]$itD.precio -lt ([double]$prodD2.precio - 0.005)) {
                            $noPermitido = "Tu usuario no puede aplicar descuentos (" + $prodD2.nombre + ")."
                            break
                        }
                    }
                }
                if ($noPermitido) {
                    Enviar-Json $context @{ ok = $false; error = $noPermitido; sinPermiso = $true } 403
                    $continuarPedido = $false
                }
            }

            if ($continuarPedido -and (-not $esAutoCli) -and (-not (Pin-Valido ([string]$data.vendedor) $data.pin))) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "PIN incorrecto para ese vendedor."; requierePin = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
                $continuarPedido = $false
            }

            # Si este mismo pedido (mismo clienteId) ya se recibio antes, no se crea
            # de nuevo: se devuelve el que ya existe, como si se acabara de crear.
            # Asi, si el telefono lo reenvia por un corte de red (crea que fallo pero
            # en realidad si llego), no se duplica la venta ni se descuenta el stock dos veces.
            $clienteIdRecibido = [string]$data.clienteId
            if ($continuarPedido -and -not [string]::IsNullOrWhiteSpace($clienteIdRecibido)) {
                $existente = $global:pedidos | Where-Object { [string]$_.clienteId -eq $clienteIdRecibido } | Select-Object -First 1
                if ($existente) {
                    Enviar-Respuesta -Context $context -Body (@{ ok = $true; id = $existente.id; repetido = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
                    $continuarPedido = $false
                }
            }

            if ($continuarPedido) {
            $items = @($data.items)

            # --- Validar stock disponible antes de aceptar el pedido ---
            $erroresStock = New-Object System.Collections.ArrayList
            foreach ($it in $items) {
                $skuItem = [string]$it.sku
                $cantidadPedida = [double]$it.cantidad
                if ([string]::IsNullOrWhiteSpace($skuItem)) { continue }
                $prod = $global:catalogo | Where-Object { $_.sku -eq $skuItem } | Select-Object -First 1
                if ($prod -and $prod.stock -ne $null -and $cantidadPedida -gt $prod.stock) {
                    [void]$erroresStock.Add("$($prod.nombre): solo quedan $($prod.stock)")
                }
            }

            if ($erroresStock.Count -gt 0) {
                $msg = "No hay stock suficiente -> " + ($erroresStock -join "; ")
                if ($esAutoCli) { $msg = "Algunos productos ya no estan disponibles en esa cantidad. Baja la cantidad o pregunta en caja." }
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = $msg } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } else {
                # Stock antes de vender (para avisar a los demas si algo se agota)
                $stockAntesPorSku = @{}
                foreach ($itA in $items) {
                    $skA = [string]$itA.sku
                    if ($skA) {
                        $pA = $global:catalogo | Where-Object { $_.sku -eq $skA } | Select-Object -First 1
                        if ($pA -and $pA.stock -ne $null) { $stockAntesPorSku[$skA] = [double]$pA.stock }
                    }
                }
                # --- Descontar stock (evita ventas negativas entre vendedores distintos) ---
                foreach ($it in $items) {
                    $skuItem = [string]$it.sku
                    if ([string]::IsNullOrWhiteSpace($skuItem)) { continue }
                    $prod = $global:catalogo | Where-Object { $_.sku -eq $skuItem } | Select-Object -First 1
                    if ($prod -and $prod.stock -ne $null) {
                        $prod.vendido = [double]$prod.vendido + [double]$it.cantidad
                        $prod.stock = [double]$prod.stockBase - [double]$prod.vendido
                        if ($prod.stock -lt 0) { $prod.stock = 0 }
                    }
                }

                $totalCalc = 0
                foreach ($it in $items) { $totalCalc += [double]$it.precio * [double]$it.cantidad }
                $totalCalc = [math]::Round($totalCalc, 2)
                # El metodo de pago se guarda desde que se crea el pedido (aunque quede
                # pendiente): asi la caja ve "Transferencia" y el total x2 a cobrar.
                # "Combinado" reparte el pedido entre efectivo y transferencia: la
                # parte en transferencia tambien se cobra x2, la de efectivo no.
                $metodoInicial = [string]$data.metodoPago
                $pagoEfectivoInicial = $null
                $pagoTransferenciaInicial = $null
                if ($metodoInicial -eq "Transferencia") {
                    $totalCobradoInicial = [math]::Round($totalCalc * 2, 2)
                } elseif ($metodoInicial -eq "Combinado") {
                    $pe = 0.0; try { $pe = [double]$data.pagoEfectivo } catch { $pe = 0.0 }
                    $pt = 0.0; try { $pt = [double]$data.pagoTransferencia } catch { $pt = 0.0 }
                    $pagoEfectivoInicial = [math]::Round($pe, 2)
                    $pagoTransferenciaInicial = [math]::Round($pt, 2)
                    $totalCobradoInicial = [math]::Round($pe + ($pt * 2), 2)
                } else {
                    $totalCobradoInicial = $totalCalc
                }

                # Si alguno de los productos venia de un pedido "para todos" que
                # este vendedor tomo, el movil manda aqui quien lo tomo y cuando,
                # para poder mostrarlo despues junto al pedido (util si dos
                # vendedores discuten quien iba a atender a un cliente).
                $origenAsignados = if ($data.origenAsignados) { @($data.origenAsignados) } else { @() }

                # Nota opcional del vendedor para la caja: solo de vendedores creados en la PC (nunca de clientes).
                $notaPedido = ""
                if ((-not $esAutoCli) -and $data.nota -and $data.vendedor -and $global:pinesVendedores.ContainsKey(([string]$data.vendedor).Trim().ToLowerInvariant()) -and (Tiene-Permiso ([string]$data.vendedor) 'mensajes')) {
                    $notaPedido = Limpiar-TextoMensaje $data.nota
                }
                $pedido = [ordered]@{
                    id                 = $global:nextId
                    clienteId          = $clienteIdRecibido
                    vendedor           = [string]$data.vendedor
                    items              = $items
                    totalProductos     = $totalCalc
                    totalCobrado       = $totalCobradoInicial
                    estado             = [string]$data.estado
                    metodoPago         = [string]$data.metodoPago
                    montoRecibido      = $data.montoRecibido
                    cambio             = $data.cambio
                    pagoEfectivo       = $pagoEfectivoInicial
                    pagoTransferencia  = $pagoTransferenciaInicial
                    hora               = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    cobradoPor         = $(if ([string]$data.estado -eq "cobrado") { "vendedor" } else { "" })
                    revisado           = $false
                    horaCobro          = $(if ([string]$data.estado -eq "cobrado") { (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") } else { "" })
                    origenAsignados    = $origenAsignados
                    nota               = $notaPedido
                }
                $global:nextId++
                [void]$global:pedidos.Add([pscustomobject]$pedido)
                Guardar-Pedidos
                try {
                    $agotadosVenta = @()
                    foreach ($skB in @($stockAntesPorSku.Keys)) {
                        $pB = $global:catalogo | Where-Object { $_.sku -eq $skB } | Select-Object -First 1
                        if ($pB -and $stockAntesPorSku[$skB] -gt 0 -and [double]$pB.stock -le 0) { $agotadosVenta += [string]$pB.nombre }
                    }
                    if ($agotadosVenta.Count -gt 0) { Avisar-A-Todos ("Se agoto: " + (Resumir-Lista $agotadosVenta 5)) "stock" ([string]$data.vendedor) }
                } catch {}

                Enviar-Respuesta -Context $context -Body (@{ ok = $true; id = $pedido.id } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }
            }   # cierre de "if ($continuarPedido)"

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/metodo$") {
            # El vendedor cambia el metodo de pago de un pedido que la caja todavia no cobro.
            $id = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }

            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            $nuevoMetodo = if ($data -and $data.metodoPago) { [string]$data.metodoPago } else { "" }
            if (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif ($pedido.estado -eq "cobrado" -or $pedido.estado -eq "cancelado") {
                $msgEstado = if ($pedido.estado -eq "cobrado") { "Este pedido ya fue cobrado y no se puede cambiar." } else { "Este pedido esta cancelado." }
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = $msgEstado; estado = [string]$pedido.estado } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif (@("Efectivo", "Transferencia", "Otro") -notcontains $nuevoMetodo) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Metodo de pago no valido." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif (-not (Pin-Valido $pedido.vendedor $data.pin)) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "PIN incorrecto."; requierePin = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
            } else {
                $pedido.metodoPago = $nuevoMetodo
                $baseTotal = [double]$pedido.totalProductos
                $pedido.totalCobrado = if ($nuevoMetodo -eq "Transferencia") { [math]::Round($baseTotal * 2, 2) } else { $baseTotal }
                Guardar-Pedidos
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/cobrar$") {
            $id = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }

            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            $metodoActual = if ($pedido -and $pedido.metodoPago) { [string]$pedido.metodoPago } else { "Efectivo" }
            if ($pedido -and ($pedido.estado -eq "cobrado" -or $pedido.estado -eq "cancelado")) {
                # Ya cobrado (o cancelado): no se vuelve a cobrar ni se cambia el metodo.
                $msgEstado = if ($pedido.estado -eq "cobrado") { "Este pedido ya fue cobrado." } else { "Este pedido esta cancelado." }
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = $msgEstado; estado = [string]$pedido.estado } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif ($pedido -and $data -and $data.metodoEsperado -and ([string]$data.metodoEsperado -ne $metodoActual)) {
                # La caja tenia en pantalla un metodo/monto y el vendedor lo cambio justo antes:
                # no se cobra a ciegas, se avisa para que revise.
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "El vendedor cambio el metodo de pago a $metodoActual. Revisa el pedido y vuelve a tocar Cobrar en Caja."; metodo = $metodoActual } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } else {
                # Cobra la CAJA si la peticion viene de la propia PC (Panel) y sin metodo
                # de pago propio; en cualquier otro caso cobro el VENDEDOR desde su movil
                # (y hace falta su PIN, si tiene uno puesto) y el pedido queda "por revisar"
                # en el Panel hasta que la caja lo confirme.
                $esLocal = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
                $cobroCaja = $esLocal -and -not ($data -and $data.metodoPago)
                if ((-not $cobroCaja) -and -not (Pin-Valido $pedido.vendedor $data.pin)) {
                    Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "PIN incorrecto."; requierePin = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
                } else {
                $metodo = if ($data -and $data.metodoPago) { [string]$data.metodoPago } elseif ($pedido.metodoPago) { [string]$pedido.metodoPago } else { "Efectivo" }
                $pedido.estado = "cobrado"
                $pedido.metodoPago = $metodo
                $baseTotal = [double]$pedido.totalProductos
                $pedido.totalCobrado = if ($metodo -eq "Transferencia") { [math]::Round($baseTotal * 2, 2) } else { $baseTotal }
                if ($data -and $null -ne $data.montoRecibido) {
                    # Se escribio cuanto dio el cliente (caja o el propio vendedor): se
                    # guarda para que el recibo calcule bien el cambio.
                    $montoRecibidoNuevo = 0.0
                    try { $montoRecibidoNuevo = [double]$data.montoRecibido } catch {}
                    $pedido | Add-Member -NotePropertyName montoRecibido -NotePropertyValue $montoRecibidoNuevo -Force
                    $pedido | Add-Member -NotePropertyName cambio -NotePropertyValue ([math]::Round($montoRecibidoNuevo - $pedido.totalCobrado, 2)) -Force
                }
                $pedido | Add-Member -NotePropertyName cobradoPor -NotePropertyValue $(if ($cobroCaja) { "caja" } else { "vendedor" }) -Force
                $pedido | Add-Member -NotePropertyName revisado -NotePropertyValue ([bool]$cobroCaja) -Force
                $pedido | Add-Member -NotePropertyName horaCobro -NotePropertyValue ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) -Force
                Guardar-Pedidos
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
                }
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/revisar$") {
            # La caja cuenta el dinero que le entrego el vendedor y confirma
            # que coincide con lo que el pedido dice que se cobro. Si no
            # coincide, el pedido se queda "por revisar" (no se marca como
            # resuelto) y se le manda una alerta al movil del vendedor.
            $id = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }
            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            if (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif ($pedido.estado -ne "cobrado") {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Solo se revisan pedidos cobrados." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } else {
                $esperado = [double]$pedido.totalCobrado
                $entregado = $esperado
                if ($data -and $null -ne $data.entregado) { try { $entregado = [double]$data.entregado } catch {} }
                $diferencia = [math]::Round($entregado - $esperado, 2)
                if ([math]::Abs($diferencia) -gt 0.009) {
                    $pedido | Add-Member -NotePropertyName discrepancia -NotePropertyValue $true -Force
                    $pedido | Add-Member -NotePropertyName montoEntregadoCaja -NotePropertyValue $entregado -Force
                    Guardar-Pedidos
                    $esperadoTxt = $esperado.ToString("0.00")
                    $entregadoTxt = $entregado.ToString("0.00")
                    $msjAlerta = "El monto del pedido #$id no coincide: tenias que entregar `$$esperadoTxt y en caja se conto `$$entregadoTxt. Revisa con la caja."
                    $alerta = [pscustomobject]@{
                        id       = $global:nextIdAlerta
                        vendedor = $pedido.vendedor
                        mensaje  = $msjAlerta
                        hora     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                        pedidoId = $id
                    }
                    $global:nextIdAlerta++
                    [void]$global:alertasVendedor.Add($alerta)
                    Guardar-Alertas
                    Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "El monto no coincide (esperado `$$esperadoTxt, contado `$$entregadoTxt). Se aviso al vendedor."; discrepancia = $true; diferencia = $diferencia } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
                } else {
                    $pedido | Add-Member -NotePropertyName revisado -NotePropertyValue $true -Force
                    $pedido | Add-Member -NotePropertyName discrepancia -NotePropertyValue $false -Force
                    $pedido | Add-Member -NotePropertyName montoEntregadoCaja -NotePropertyValue $entregado -Force
                    Guardar-Pedidos
                    Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
                }
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/problema$") {
            # La caja reporta que el vendedor no le entrego lo que corresponde
            # (dinero, un producto, etc.) al cobrar un pedido. Queda una nota
            # visible en la tarjeta y se le manda un aviso a su movil.
            $id = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }
            $motivo = if ($data -and $data.motivo) { ([string]$data.motivo).Trim() } else { "" }
            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            if (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } else {
                $pedido | Add-Member -NotePropertyName problemaReportado -NotePropertyValue $true -Force
                $pedido | Add-Member -NotePropertyName problemaMotivo -NotePropertyValue $motivo -Force
                $pedido | Add-Member -NotePropertyName horaProblema -NotePropertyValue ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) -Force
                Guardar-Pedidos
                $msjAlerta = "La caja reporto un problema con el pedido #$id" + $(if ($motivo) { ": $motivo" } else { "." }) + " Hablen antes de seguir."
                $alerta = [pscustomobject]@{
                    id       = $global:nextIdAlerta
                    vendedor = $pedido.vendedor
                    mensaje  = $msjAlerta
                    hora     = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    pedidoId = $id
                }
                $global:nextIdAlerta++
                [void]$global:alertasVendedor.Add($alerta)
                Guardar-Alertas
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "GET" -and $path -eq "/etiquetas") {
            Enviar-Respuesta -Context $context -Body (Inyectar-Guardian $htmlEtiquetas)

        } elseif ($method -eq "GET" -and $path -eq "/api/catalogo/nuevos") {
            # Productos que aparecieron por primera vez en el ultimo Excel cargado (para sus etiquetas).
            Enviar-Json $context @{ ok = $true; fecha = [string]$global:productosNuevos.fecha; skus = @($global:productosNuevos.skus) } 200 4

        } elseif ($method -eq "GET" -and $path -eq "/api/pos/estado") {
            # El movil pregunta si el administrador ya puso la clave y si este movil sigue con el modo activado.
            $vPosE = Pos-VendedorDe $request
            Enviar-Json $context @{ ok = $true; claveDefinida = (-not [string]::IsNullOrEmpty($global:posClaveHash)); activo = (Pos-Token-Valido $request $vPosE) }

        } elseif ($method -eq "POST" -and $path -eq "/api/pos/clave") {
            # Solo desde la PC: pone, cambia o quita la clave de administrador del modo punto de venta.
            # Al cambiarla o quitarla, todos los moviles pierden el modo y deben volver a activarlo.
            $dPosC = Leer-CuerpoJson $request
            $localPosC = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $clavePosN = if ($dPosC -and $dPosC.clave) { ([string]$dPosC.clave).Trim() } else { "" }
            if (-not $localPosC) {
                Enviar-Json $context @{ ok = $false; error = "La clave de administrador solo se cambia desde la PC." } 403
            } elseif ($clavePosN -and (($clavePosN.Length -lt 4) -or ($clavePosN.Length -gt 20))) {
                Enviar-Json $context @{ ok = $false; error = "La clave debe tener de 4 a 20 caracteres." } 400
            } else {
                $global:posClaveHash = if ($clavePosN) { Pos-Hash $clavePosN } else { "" }
                $global:posTokens = @{}
                $global:posIntentos = @{}
                Guardar-Pos
                Enviar-Json $context @{ ok = $true; definida = ($clavePosN -ne "") }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/pos/activar") {
            # El movil manda la clave de administrador; si es correcta, recibe un token propio.
            $dPosA = Leer-CuerpoJson $request
            $vPosA = if ($dPosA -and $dPosA.vendedor) { ([string]$dPosA.vendedor).Trim() } else { "" }
            $ipPosA = [string]$request.RemoteEndPoint.Address
            $esperaPosA = Pos-Bloqueado $ipPosA
            if ($esperaPosA -gt 0) {
                Enviar-Json $context @{ ok = $false; error = "Demasiados intentos. Espera $esperaPosA segundos." } 429
            } elseif ([string]::IsNullOrEmpty($global:posClaveHash)) {
                Enviar-Json $context @{ ok = $false; error = "El administrador todavia no puso la clave en la PC (menu, PIN de vendedores)."; sinClave = $true } 409
            } elseif ((-not $vPosA) -or (-not $global:pinesVendedores.ContainsKey($vPosA.ToLowerInvariant()))) {
                Enviar-Json $context @{ ok = $false; error = "Entra primero como vendedor." } 403
            } elseif (-not (Pin-Valido $vPosA $dPosA.pin)) {
                Enviar-Json $context @{ ok = $false; error = "PIN incorrecto."; requierePin = $true } 403
            } elseif ((Pos-Hash ([string]$dPosA.clave)) -ne $global:posClaveHash) {
                Pos-Fallo $ipPosA
                Enviar-Json $context @{ ok = $false; error = "Clave de administrador incorrecta." } 403
            } else {
                if ($global:posIntentos.ContainsKey($ipPosA)) { $global:posIntentos.Remove($ipPosA) }
                $tokPos = [guid]::NewGuid().ToString("N") + [guid]::NewGuid().ToString("N")
                $global:posTokens[(Pos-Hash $tokPos)] = $vPosA.ToLowerInvariant()
                Guardar-Pos
                Enviar-Json $context @{ ok = $true; token = $tokPos }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/pos/ventas") {
            # Historial de ventas cobradas + reporte de efectivo / transferencia (solo moviles con el modo activado).
            $vPosV = Pos-VendedorDe $request
            if (-not (Pos-Token-Valido $request $vPosV)) {
                Enviar-Json $context @{ ok = $false; error = "El modo punto de venta no esta activo en este telefono."; posInactivo = $true } 403
            } else {
                $diasPV = 1
                if ($request.QueryString["dias"]) { try { $diasPV = [int]$request.QueryString["dias"] } catch { $diasPV = 1 } }
                if ($diasPV -lt 1) { $diasPV = 1 }
                if ($diasPV -gt 90) { $diasPV = 90 }
                $quienPV = if ($request.QueryString["quien"] -eq "yo") { $vPosV.ToLowerInvariant() } else { "" }
                try {
                    Enviar-Json $context (Calcular-PosVentas $diasPV $quienPV) 200 8
                } catch {
                    Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/pos/devolucion") {
            # Devolucion desde el movil en modo punto de venta (valida contra la venta original).
            $dPosD = Leer-CuerpoJson $request
            $vPosD = Pos-VendedorDe $request
            if (-not (Pos-Token-Valido $request $vPosD)) {
                Enviar-Json $context @{ ok = $false; error = "El modo punto de venta no esta activo en este telefono."; posInactivo = $true } 403
            } elseif (-not (Pin-Valido $vPosD $dPosD.pin)) {
                Enviar-Json $context @{ ok = $false; error = "PIN incorrecto."; requierePin = $true } 403
            } else {
                try {
                    $resPosD = Registrar-DevolucionPos $dPosD $vPosD
                    Enviar-Json $context $resPosD.cuerpo ([int]$resPosD.codigo) 8
                } catch {
                    Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/mensajes") {
            # Mensaje corto del movil de un vendedor hacia la PC (sin pedido). Hace falta ser un
            # vendedor creado en la PC y su PIN: un cliente (autoservicio) no puede mandar mensajes.
            $dMsg = Leer-CuerpoJson $request
            $vMsg = if ($dMsg -and $dMsg.vendedor) { ([string]$dMsg.vendedor).Trim() } else { "" }
            $txtMsg = if ($dMsg) { Limpiar-TextoMensaje $dMsg.texto } else { "" }
            if ((-not $vMsg) -or (-not $global:pinesVendedores.ContainsKey($vMsg.ToLowerInvariant())) -or (-not (Pin-Valido $vMsg $dMsg.pin))) {
                Enviar-Json $context @{ ok = $false; error = "PIN incorrecto."; requierePin = $true } 403
            } elseif (-not (Tiene-Permiso $vMsg 'mensajes')) {
                Enviar-Json $context @{ ok = $false; error = "Tu usuario no puede enviar mensajes. Pidele al encargado que lo active desde la PC."; sinPermiso = $true } 403
            } elseif (-not $txtMsg) {
                Enviar-Json $context @{ ok = $false; error = "Escribe el mensaje." } 400
            } else {
                $nombreMsg = [string]$global:pinesVendedores[$vMsg.ToLowerInvariant()].nombre
                $regMsg = [pscustomobject]@{ id = $global:nextIdMensaje; vendedor = $nombreMsg; texto = $txtMsg; hora = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); leido = $false }
                $global:nextIdMensaje++
                [void]$global:mensajesPc.Add($regMsg)
                Guardar-MensajesPc
                Enviar-Json $context @{ ok = $true; id = $regMsg.id }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/mensajes/pc") {
            # Solo la PC: entrega (una sola vez) los mensajes de vendedores que todavia no se han mostrado.
            if (-not [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)) {
                Enviar-Json $context @{ ok = $false; error = "Solo desde la PC." } 403
            } else {
                $pendMsg = @($global:mensajesPc | Where-Object { -not $_.leido })
                if ($pendMsg.Count -gt 0) {
                    foreach ($mP in $pendMsg) { $mP.leido = $true }
                    Guardar-MensajesPc
                }
                Enviar-Json $context @{ ok = $true; mensajes = @($pendMsg) } 200 4
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/mensajes/enviar") {
            # Solo la PC: mensaje corto para un vendedor (o para todos). Le llega como aviso en el movil.
            $dMsgE = Leer-CuerpoJson $request
            $txtMsgE = if ($dMsgE) { Limpiar-TextoMensaje $dMsgE.texto } else { "" }
            $paraTodosMsg = [bool]($dMsgE -and $dMsgE.todos)
            $destMsg = if ($dMsgE -and $dMsgE.vendedor) { ([string]$dMsgE.vendedor).Trim() } else { "" }
            if (-not [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)) {
                Enviar-Json $context @{ ok = $false; error = "Los mensajes a vendedores se envian desde la PC." } 403
            } elseif (-not $txtMsgE) {
                Enviar-Json $context @{ ok = $false; error = "Escribe el mensaje." } 400
            } elseif ($paraTodosMsg) {
                Avisar-A-Todos ("Mensaje de la caja: " + $txtMsgE) "mensaje"
                Enviar-Json $context @{ ok = $true }
            } elseif ($destMsg -and $global:pinesVendedores.ContainsKey($destMsg.ToLowerInvariant())) {
                Agregar-AlertaVendedor ([string]$global:pinesVendedores[$destMsg.ToLowerInvariant()].nombre) ("Mensaje de la caja: " + $txtMsgE) $null "mensaje"
                Enviar-Json $context @{ ok = $true }
            } else {
                Enviar-Json $context @{ ok = $false; error = "Ese vendedor no existe." } 404
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/permisos") {
            # La app movil pregunta que puede hacer este vendedor.
            $nv = if ($request.QueryString["vendedor"]) { $request.QueryString["vendedor"].Trim() } else { "" }
            Enviar-Json $context @{ ok = $true; permisos = (Obtener-Permisos $nv) }

        } elseif ($method -eq "GET" -and $path -eq "/api/permisos/todos") {
            # El Panel pide todos los vendedores con sus permisos actuales.
            $listaPerm = @()
            foreach ($k in @($global:pinesVendedores.Keys | Sort-Object)) {
                $nom = [string]$global:pinesVendedores[$k].nombre
                $listaPerm += [pscustomobject]@{ nombre = $nom; permisos = (Obtener-Permisos $nom) }
            }
            $clavesPerm = @($global:catalogoPermisos | ForEach-Object { [pscustomobject]@{ k = $_.k; t = $_.t; d = $_.d } })
            Enviar-Json $context @{ ok = $true; claves = $clavesPerm; vendedores = @($listaPerm) } 200 8

        } elseif ($method -eq "POST" -and $path -eq "/api/permisos") {
            # Solo desde la PC. Guarda los permisos de un vendedor y le avisa al movil.
            $dataPerm = Leer-CuerpoJson $request
            $esLocalPerm = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $vendPerm = if ($dataPerm -and $dataPerm.vendedor) { ([string]$dataPerm.vendedor).Trim() } else { "" }
            $clavePerm = $vendPerm.ToLowerInvariant()
            if (-not $esLocalPerm) {
                Enviar-Json $context @{ ok = $false; error = "Los permisos solo se cambian desde la PC." } 403
            } elseif ((-not $vendPerm) -or (-not $global:pinesVendedores.ContainsKey($clavePerm))) {
                Enviar-Json $context @{ ok = $false; error = "Ese vendedor no existe." } 404
            } elseif (-not $dataPerm.permisos) {
                Enviar-Json $context @{ ok = $false; error = "Faltan los permisos." } 400
            } else {
                $actualPerm = Obtener-Permisos $vendPerm
                $cambiosPerm = New-Object System.Collections.ArrayList
                foreach ($c in $global:catalogoPermisos) {
                    if ($dataPerm.permisos.PSObject.Properties.Name -contains $c.k) {
                        $nuevoV = [bool]$dataPerm.permisos.($c.k)
                        if ([bool]$actualPerm[$c.k] -ne $nuevoV) {
                            $actualPerm[$c.k] = $nuevoV
                            [void]$cambiosPerm.Add($c.t + $(if ($nuevoV) { ": SI" } else { ": NO" }))
                        }
                    }
                }
                $guardarPerm = @{}
                foreach ($c in $global:catalogoPermisos) { $guardarPerm[$c.k] = [bool]$actualPerm[$c.k] }
                $global:permisosVendedores[$clavePerm] = $guardarPerm
                Guardar-Permisos
                if ($cambiosPerm.Count -gt 0) {
                    Agregar-AlertaVendedor $vendPerm ("La caja cambio tus permisos: " + (Resumir-Lista $cambiosPerm 4)) $null "permisos"
                }
                Enviar-Json $context @{ ok = $true; permisos = $actualPerm }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/reabastecer") {
            # Productos por debajo de su minimo de seguridad + cantidad sugerida a comprar.
            $itemsReab = @($global:catalogo |
                Where-Object { $mm = Minimo-Efectivo $_; $_.stock -ne $null -and $mm -gt 0 -and [double]$_.stock -le $mm } |
                Sort-Object -Property @{ Expression = { [double]$_.stock } } |
                ForEach-Object {
                    $mm = Minimo-Efectivo $_
                    $sug = [math]::Ceiling([double](($mm * 2) - [double]$_.stock))
                    if ($sug -lt 1) { $sug = 1 }
                    $skuR = [string]$_.sku
                    [pscustomobject]@{
                        sku = $skuR; nombre = $_.nombre; stock = $_.stock; minimo = $mm; sugerido = $sug
                        personalizado = [bool]($skuR -and $global:minimosStock.ContainsKey($skuR))
                    }
                })
            Enviar-Json $context @{ ok = $true; umbral = [int]$global:configApp.umbralStockBajo; items = $itemsReab } 200 5

        } elseif ($method -eq "POST" -and $path -eq "/api/minimos") {
            # Fija (o quita) el minimo de seguridad de un producto. Solo desde la PC.
            $dataMin = Leer-CuerpoJson $request
            $esLocalMin = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $skuMin = if ($dataMin -and $dataMin.sku) { ([string]$dataMin.sku).Trim() } else { "" }
            if (-not $esLocalMin) {
                Enviar-Json $context @{ ok = $false; error = "Solo se puede cambiar desde la PC." } 403
            } elseif (-not $skuMin) {
                Enviar-Json $context @{ ok = $false; error = "Falta el SKU." } 400
            } else {
                $valMin = $null
                if ($dataMin.PSObject.Properties.Name -contains 'minimo' -and $dataMin.minimo -ne $null -and "$($dataMin.minimo)" -ne "") {
                    $tmp = 0.0
                    if ([double]::TryParse([string]$dataMin.minimo, [ref]$tmp) -and $tmp -ge 0) { $valMin = $tmp }
                }
                if ($valMin -eq $null) { $global:minimosStock.Remove($skuMin) } else { $global:minimosStock[$skuMin] = $valMin }
                Guardar-Minimos
                Enviar-Json $context @{ ok = $true }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/reabastecer/imprimir") {
            # Imprime la lista de compras en el ticket de 32 columnas.
            $dataCmp = Leer-CuerpoJson $request
            $esLocalCmp = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if (-not $esLocalCmp) {
                Enviar-Json $context @{ ok = $false; error = "Solo se puede imprimir desde la PC." } 403
            } elseif ((-not $dataCmp) -or @($dataCmp.items).Count -eq 0) {
                Enviar-Json $context @{ ok = $false; error = "No hay articulos en la lista." } 400
            } else {
                try {
                    $okCmp = Imprimir-Texto (Generar-TextoListaCompras @($dataCmp.items))
                    if ($okCmp) { Enviar-Json $context @{ ok = $true } } else { Enviar-Json $context @{ ok = $false; error = "Windows no confirmo la impresion." } 500 }
                } catch {
                    Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/etiquetas/imprimir") {
            # Imprime etiquetas en la misma termica del ticket (RAW / ESC-POS). Solo desde la PC.
            $dataEt = Leer-CuerpoJson $request
            $esLocalEt = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $totalEt = 0
            if ($dataEt) { foreach ($x in @($dataEt.items)) { try { $totalEt += [int]$x.copias } catch {} } }
            if (-not $esLocalEt) {
                Enviar-Json $context @{ ok = $false; error = "Las etiquetas se imprimen desde la PC." } 403
            } elseif ($totalEt -lt 1) {
                Enviar-Json $context @{ ok = $false; error = "No hay etiquetas para imprimir." } 400
            } elseif ($totalEt -gt 400) {
                Enviar-Json $context @{ ok = $false; error = "Maximo 400 etiquetas por trabajo (hay $totalEt)." } 400
            } else {
                try {
                    $bytesEt = Generar-BytesEtiquetas @($dataEt.items) $dataEt.opciones
                    if ([RawPrinterHelper]::EnviarBytes($nombreImpresora, $bytesEt)) { Enviar-Json $context @{ ok = $true; etiquetas = $totalEt } }
                    else { Enviar-Json $context @{ ok = $false; error = "Windows no confirmo la impresion." } 500 }
                } catch {
                    Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                }
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/etiquetas/raw") {
            # Recibe los bytes ESC/POS ya dibujados por la pagina de etiquetas y los manda
            # tal cual a la termica (RAW). Reemplaza al "ayudante de impresion" (puerto 8787).
            $esLocalRaw = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if (-not $esLocalRaw) {
                Enviar-Json $context @{ ok = $false; error = "Solo se imprime desde la PC." } 403
            } else {
                $msRaw = New-Object System.IO.MemoryStream
                $request.InputStream.CopyTo($msRaw)
                $bytesRaw = $msRaw.ToArray()
                if ($bytesRaw.Length -lt 2) {
                    Enviar-Json $context @{ ok = $false; error = "No llegaron datos para imprimir." } 400
                } elseif ($bytesRaw.Length -gt 30000000) {
                    Enviar-Json $context @{ ok = $false; error = "El trabajo es demasiado grande." } 413
                } else {
                    try {
                        if ([RawPrinterHelper]::EnviarBytes($nombreImpresora, $bytesRaw)) { Enviar-Json $context @{ ok = $true; bytes = $bytesRaw.Length } }
                        else { Enviar-Json $context @{ ok = $false; error = "Windows no confirmo la impresion." } 500 }
                    } catch {
                        Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                    }
                }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/red") {
            $rr = Obtener-RedActiva
            Enviar-Json $context @{ ok = $true; red = $rr; config = $global:redConfig; puerto = $port } 5

        } elseif ($method -eq "POST" -and $path -eq "/api/red/fijar") {
            $dataRed = Leer-CuerpoJson $request
            $esLocalRed = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if (-not $esLocalRed) {
                Enviar-Json $context @{ ok = $false; error = "Solo se cambia desde la PC." } 403
            } else {
                $oct = 0; try { $oct = [int]$dataRed.octeto } catch { $oct = 0 }
                $resRed = Aplicar-IpFija $oct
                if ($resRed.ok) {
                    $global:redConfig.ultimoOcteto = $oct
                    $global:redConfig.auto = [bool]$dataRed.auto
                    Guardar-RedConfig
                }
                Enviar-Json $context $resRed $(if ($resRed.ok) { 200 } else { 409 })
            }

        } elseif ($method -eq "POST" -and $path -eq "/api/red/dhcp") {
            $esLocalRed2 = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if (-not $esLocalRed2) {
                Enviar-Json $context @{ ok = $false; error = "Solo se cambia desde la PC." } 403
            } else {
                $global:redConfig.auto = $false
                Guardar-RedConfig
                $resRed2 = Volver-IpAutomatica
                Enviar-Json $context $resRed2 $(if ($resRed2.ok) { 200 } else { 500 })
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/devoluciones") {
            $ultDev = @($global:devoluciones | Select-Object -Last 30)
            [array]::Reverse($ultDev)
            Enviar-Json $context @{ ok = $true; devoluciones = $ultDev } 200 8

        } elseif ($method -eq "POST" -and $path -eq "/api/devoluciones") {
            # Registra una devolucion / cambio / garantia, (opcional) reintegra el
            # stock y imprime el comprobante de 32 columnas. Solo desde la PC.
            $dataDev = Leer-CuerpoJson $request
            $esLocalDev = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $itemsDev = @()
            if ($dataDev) {
                foreach ($it in @($dataDev.items)) {
                    $cantD = 0.0; try { $cantD = [double]$it.cantidad } catch { $cantD = 0.0 }
                    if ($cantD -le 0) { continue }
                    $prD = 0.0; try { $prD = [double]$it.precio } catch { $prD = 0.0 }
                    $itemsDev += [pscustomobject]@{ sku = [string]$it.sku; nombre = [string]$it.nombre; cantidad = $cantD; precio = $prD }
                }
            }
            if (-not $esLocalDev) {
                Enviar-Json $context @{ ok = $false; error = "Las devoluciones se registran desde la PC." } 403
            } elseif ($itemsDev.Count -eq 0) {
                Enviar-Json $context @{ ok = $false; error = "Agrega al menos un producto." } 400
            } else {
                $tipoDev = if (@("devolucion", "cambio", "garantia") -contains [string]$dataDev.tipo) { [string]$dataDev.tipo } else { "devolucion" }
                $reintegrar = [bool]$dataDev.reintegrarStock
                if ($reintegrar) {
                    foreach ($it in $itemsDev) {
                        if ([string]::IsNullOrWhiteSpace($it.sku)) { continue }
                        $prodD = $global:catalogo | Where-Object { $_.sku -eq $it.sku } | Select-Object -First 1
                        if ($prodD -and $prodD.stockBase -ne $null) {
                            $prodD.vendido = [double]$prodD.vendido - [double]$it.cantidad
                            if ($prodD.vendido -lt 0) { $prodD.vendido = 0 }
                            $prodD.stock = [double]$prodD.stockBase - [double]$prodD.vendido
                            if ($prodD.stock -lt 0) { $prodD.stock = 0 }
                        }
                    }
                }
                $totalDev = 0.0
                foreach ($it in $itemsDev) { $totalDev += [double]$it.precio * [double]$it.cantidad }
                $difDev = $null
                if ($dataDev.diferencia -ne $null -and "$($dataDev.diferencia)" -ne "") { try { $difDev = [double]$dataDev.diferencia } catch { $difDev = $null } }
                $registroDev = [pscustomobject]@{
                    id               = $global:nextIdDevolucion
                    tipo             = $tipoDev
                    hora             = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    pedidoId         = $(if ($dataDev.pedidoId) { [string]$dataDev.pedidoId } else { "" })
                    cliente          = ([string]$dataDev.cliente).Trim()
                    atendio          = ([string]$dataDev.atendio).Trim()
                    motivo           = ([string]$dataDev.motivo).Trim()
                    items            = $itemsDev
                    total            = [math]::Round($totalDev, 2)
                    metodoReembolso  = ([string]$dataDev.metodoReembolso).Trim()
                    entregado        = ([string]$dataDev.entregado).Trim()
                    diferencia       = $difDev
                    reintegrarStock  = $reintegrar
                }
                $global:nextIdDevolucion++
                [void]$global:devoluciones.Add($registroDev)
                Guardar-Devoluciones
                $impresoDev = $false
                $errorDev = ""
                try { $impresoDev = [bool](Imprimir-Texto (Generar-TextoDevolucion $registroDev)) } catch { $errorDev = "$($_.Exception.Message)" }
                Enviar-Json $context @{ ok = $true; id = $registroDev.id; impreso = $impresoDev; errorImpresion = $errorDev }
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/devoluciones/(\d+)/imprimir$") {
            $idDevR = [int]$Matches[1]
            $devR = $global:devoluciones | Where-Object { [int]$_.id -eq $idDevR } | Select-Object -First 1
            if (-not $devR) {
                Enviar-Json $context @{ ok = $false; error = "Devolucion no encontrada." } 404
            } else {
                try {
                    if (Imprimir-Texto (Generar-TextoDevolucion $devR)) { Enviar-Json $context @{ ok = $true } } else { Enviar-Json $context @{ ok = $false; error = "Windows no confirmo la impresion." } 500 }
                } catch {
                    Enviar-Json $context @{ ok = $false; error = "$($_.Exception.Message)" } 500
                }
            }

        } elseif ($method -eq "GET" -and $path -eq "/api/alertas") {
            # El movil del vendedor pregunta si tiene alguna alerta nueva
            # (por ejemplo, un desajuste de dinero al revisar en caja).
            $nombreVend = if ($request.QueryString["vendedor"]) { $request.QueryString["vendedor"].Trim() } else { "" }
            $paraEste = @($global:alertasVendedor | Where-Object { $_.vendedor -eq $nombreVend })
            if ($paraEste.Count -gt 0) {
                foreach ($a in $paraEste) { [void]$global:alertasVendedor.Remove($a) }
                Guardar-Alertas
            }
            Enviar-Respuesta -Context $context -Body (@{ ok = $true; alertas = $paraEste } | ConvertTo-Json -Depth 5) -ContentType "application/json; charset=utf-8"

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/cancelar$") {
            $id = [int]$Matches[1]
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }
            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            $esLocalCanc = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            if (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif ($pedido.estado -eq "cancelado") {
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            } elseif ($pedido.estado -eq "cobrado") {
                # Una vez cobrado (desde el movil o la PC) ya no se puede anular ni editar.
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Este pedido ya fue cobrado y no se puede anular."; estado = "cobrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif ((-not $esLocalCanc) -and -not (Pin-Valido $pedido.vendedor ($data.pin))) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "PIN incorrecto."; requierePin = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
            } else {
                # Devolver al stock las unidades de este pedido (si tenian control de stock)
                foreach ($it in @($pedido.items)) {
                    $skuItem = [string]$it.sku
                    if ([string]::IsNullOrWhiteSpace($skuItem)) { continue }
                    $prod = $global:catalogo | Where-Object { $_.sku -eq $skuItem } | Select-Object -First 1
                    if ($prod -and $prod.stockBase -ne $null) {
                        $prod.vendido = [double]$prod.vendido - [double]$it.cantidad
                        if ($prod.vendido -lt 0) { $prod.vendido = 0 }
                        $prod.stock = [double]$prod.stockBase - [double]$prod.vendido
                        if ($prod.stock -lt 0) { $prod.stock = 0 }
                    }
                }
                $pedido.estado = "cancelado"
                Guardar-Pedidos
                if ($esLocalCanc -and $pedido.vendedor) {
                    Agregar-AlertaVendedor ([string]$pedido.vendedor) ("La caja anulo tu pedido #" + $pedido.id + ". El stock se devolvio.") $pedido.id "anulado"
                }
                Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/agregar$") {
            # La caja agrega productos a un pedido pendiente ya recibido (ej.
            # algo que el cliente pidio despues de mandar el pedido). Solo
            # desde la PC. Se suma al pedido, se recalcula el total y se le
            # avisa al vendedor en su movil (alerta + su lista se actualiza sola).
            $id = [int]$Matches[1]
            $esLocalAgregar = [System.Net.IPAddress]::IsLoopback($request.RemoteEndPoint.Address)
            $reader = New-Object System.IO.StreamReader($request.InputStream, [System.Text.Encoding]::UTF8)
            $bodyText = $reader.ReadToEnd()
            $reader.Close()
            $data = $null
            if ($bodyText -and $bodyText.Trim().Length -gt 0) { $data = $bodyText | ConvertFrom-Json }
            $itemsNuevos = @($data.items)
            $claveEnvioAgr = if ($data -and $data.claveEnvio) { "agr|$id|" + [string]$data.claveEnvio } else { "" }
            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            if (-not $esLocalAgregar) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Solo se puede agregar productos desde la PC." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 403
            } elseif (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } elseif ($pedido.estado -ne "pendiente") {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Este pedido ya no esta pendiente." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif ($pedido.metodoPago -eq "Combinado") {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "No se puede agregar productos a un pedido con pago Combinado; editalo desde el movil del vendedor." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
            } elseif ($itemsNuevos.Count -eq 0) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Agrega al menos un producto." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 400
            } elseif ($respPreviaAgr = (Clave-EnvioBuscar $claveEnvioAgr)) {
                # Mismos productos ya agregados con este mismo envio: no se suman dos veces.
                Enviar-Respuesta -Context $context -Body $respPreviaAgr -ContentType "application/json; charset=utf-8"
            } else {
                # --- Validar stock antes de agregar (igual que al crear un pedido) ---
                $erroresStockAg = New-Object System.Collections.ArrayList
                foreach ($it in $itemsNuevos) {
                    $skuItem = [string]$it.sku
                    $cantidadPedida = [double]$it.cantidad
                    if ([string]::IsNullOrWhiteSpace($skuItem)) { continue }
                    $prod = $global:catalogo | Where-Object { $_.sku -eq $skuItem } | Select-Object -First 1
                    if ($prod -and $prod.stock -ne $null -and $cantidadPedida -gt $prod.stock) {
                        [void]$erroresStockAg.Add("$($prod.nombre): solo quedan $($prod.stock)")
                    }
                }
                if ($erroresStockAg.Count -gt 0) {
                    $msgAg = "No hay stock suficiente -> " + ($erroresStockAg -join "; ")
                    Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = $msgAg } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 409
                } else {
                    $itemsActuales = New-Object System.Collections.ArrayList
                    foreach ($it in @($pedido.items)) { [void]$itemsActuales.Add($it) }
                    $descripcionAgregado = New-Object System.Collections.ArrayList
                    foreach ($it in $itemsNuevos) {
                        $skuItem = [string]$it.sku
                        $cantidad = [double]$it.cantidad
                        if (-not [string]::IsNullOrWhiteSpace($skuItem)) {
                            $prod = $global:catalogo | Where-Object { $_.sku -eq $skuItem } | Select-Object -First 1
                            if ($prod -and $prod.stock -ne $null) {
                                $prod.vendido = [double]$prod.vendido + $cantidad
                                $prod.stock = [double]$prod.stockBase - [double]$prod.vendido
                                if ($prod.stock -lt 0) { $prod.stock = 0 }
                            }
                        }
                        $existenteItem = $itemsActuales | Where-Object { $skuItem -ne "" -and [string]$_.sku -eq $skuItem } | Select-Object -First 1
                        if ($existenteItem) {
                            $existenteItem.cantidad = [double]$existenteItem.cantidad + $cantidad
                        } else {
                            [void]$itemsActuales.Add([pscustomobject]@{ sku = $skuItem; nombre = [string]$it.nombre; precio = [double]$it.precio; cantidad = $cantidad })
                        }
                        [void]$descripcionAgregado.Add("$cantidad x $($it.nombre)")
                    }
                    $pedido.items = @($itemsActuales)
                    $nuevoTotalProductos = 0
                    foreach ($it in $pedido.items) { $nuevoTotalProductos += [double]$it.precio * [double]$it.cantidad }
                    $nuevoTotalProductos = [math]::Round($nuevoTotalProductos, 2)
                    $pedido.totalProductos = $nuevoTotalProductos
                    $pedido.totalCobrado = if ($pedido.metodoPago -eq "Transferencia") { [math]::Round($nuevoTotalProductos * 2, 2) } else { $nuevoTotalProductos }
                    Guardar-Pedidos

                    $msjAlertaAg = "La caja MODIFICO tu pedido #$id" + ": agrego " + ($descripcionAgregado -join ", ") + ". Nuevo total a cobrar: $" + (Formato-Monto $pedido.totalCobrado) + ". Revisalo antes de cobrar."
                    Agregar-AlertaVendedor ([string]$pedido.vendedor) $msjAlertaAg $id "pedido"

                    Clave-EnvioGuardar $claveEnvioAgr (@{ ok = $true; repetido = $true } | ConvertTo-Json -Compress)
                    Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
                }
            }

        } elseif ($method -eq "POST" -and $path -match "^/api/pedidos/(\d+)/imprimir$") {
            $id = [int]$Matches[1]
            $pedido = $global:pedidos | Where-Object { [int]$_.id -eq $id }
            if (-not $pedido) {
                Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Pedido no encontrado" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 404
            } else {
                try {
                    $texto = Generar-TextoRecibo $pedido
                    $bytesRecibo = [System.Text.Encoding]::ASCII.GetBytes($texto)
                    $exito = [RawPrinterHelper]::EnviarBytes($nombreImpresora, $bytesRecibo)
                    if ($exito) {
                        Enviar-Respuesta -Context $context -Body (@{ ok = $true } | ConvertTo-Json) -ContentType "application/json; charset=utf-8"
                    } else {
                        Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "Windows no confirmo la impresion." } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
                    }
                } catch {
                    Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "$($_.Exception.Message)" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
                }
            }

        } else {
            Enviar-Respuesta -Context $context -Body "No encontrado" -ContentType "text/plain; charset=utf-8" -StatusCode 404
        }
    } catch {
        Write-Host "Error procesando solicitud: $_"
        try {
            Enviar-Respuesta -Context $context -Body (@{ ok = $false; error = "$_" } | ConvertTo-Json) -ContentType "application/json; charset=utf-8" -StatusCode 500
        } catch {}
    }
}

$listener.Stop()
$listener.Close()