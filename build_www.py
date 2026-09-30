# Prepara www/index.html (la app del vendedor para la APK) a partir de servidor_pedidos.ps1
import re, os
s = open('servidor_pedidos.ps1', encoding='utf-8-sig').read()
h = re.search(r"\$htmlVendedor = @'\r?\n(.*?)\r?\n'@", s, re.S).group(1)
g = re.search(r"\$global:licGuardJs = @'\r?\n(.*?)\r?\n'@", s, re.S).group(1)

setup = r"""
<script>
/* --- APK: direccion de la PC. Todas las rutas /api, /foto... se piden a la PC --- */
(function () {
  var K = 'ttPcUrl', u = '';
  try { u = localStorage.getItem(K) || ''; } catch (e) {}
  window.ttPedirPc = function () {
    var v = prompt('Direccion de la PC (la que muestra el servidor, ej. 192.168.43.55):', u.replace(/^https?:\/\//, ''));
    if (v === null) return;
    v = v.trim().replace(/\/+$/, '').replace(/\/vendedor.*$/, '');
    if (!v) return;
    if (!/^https?:\/\//.test(v)) v = 'http://' + v;
    if (!/:\d+$/.test(v.replace(/^https?:\/\//, ''))) v += ':8080';
    try { localStorage.setItem(K, v); } catch (e) {}
    location.reload();
  };
  if (!u) { window.ttPedirPc(); try { u = localStorage.getItem(K) || ''; } catch (e) {} }
  if (u) document.write('<base href="' + u + '/">');
  window.addEventListener('DOMContentLoaded', function () {
    var caja = document.getElementById('secRedSenal');
    if (!caja) return;
    var b = document.createElement('button');
    b.className = 'btn btn-secundario'; b.style.marginTop = '8px';
    b.textContent = 'Cambiar direccion de la PC' + (u ? ' (' + u.replace(/^https?:\/\//, '') + ')' : '');
    b.onclick = function () { window.ttPedirPc(); };
    caja.appendChild(b);
  });
})();
</script>
"""
i = h.index('<head>') + 6
h = h[:i] + setup + g + h[i:]
os.makedirs('www', exist_ok=True)
open('www/index.html', 'w', encoding='utf-8').write(h)
print('www/index.html listo:', len(h), 'bytes')
