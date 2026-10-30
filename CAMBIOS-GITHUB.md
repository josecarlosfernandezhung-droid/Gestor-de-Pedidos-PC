# Cambios para Gestor-de-Pedidos-PC

Repo: https://github.com/josecarlosfernandezhung-droid/Gestor-de-Pedidos-PC

## 1. Sustituir / añadir en la raíz del repo

| Archivo | Acción |
|---------|--------|
| `instalador.iss` | **Reemplazar** por el nuevo (publisher + licencia obligatoria) |
| `LICENCIA_INSTALACION.txt` | **Añadir** (Términos + Privacidad) |
| `servidor_pedidos.ps1` | **Reemplazar** por la versión corregida (QR, menú, etc.) |

## 2. Qué cambia el instalador

Antes:
- AppName / Publisher: Toto Tools
- Sin pantalla de licencia

Ahora:
- Publisher: **José Carlos Fernández**
- Nombre: **Servidor de Pedidos**
- Pantalla de **Términos y Política de privacidad**
- El usuario **debe aceptar** para continuar (Inno Setup con LicenseFile)

## 3. Workflow (instalador.yml)

**No hace falta cambiarlo.** Sigue así:

```yaml
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" instalador.iss
```

Solo asegúrate de que `LICENCIA_INSTALACION.txt` esté en la **raíz** del repo (junto a `instalador.iss`).

## 4. Cómo publicar

1. Sube los archivos a `main`
2. Actions → **Construir instalador PC** → Run workflow
3. Descarga el artifact `Instalar-ServidorPedidos`
4. Al ejecutar el Setup verás: Bienvenida → **Licencia (aceptar)** → Instalar

## 5. App móvil (QR)

El `servidor_pedidos.ps1` corregido usa el mismo escáner que los códigos de barras (`TotoBg.scan()`). Tras recompilar el .exe e instalar, recarga la app del vendedor.
