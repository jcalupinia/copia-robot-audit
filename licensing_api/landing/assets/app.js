// JS mínimo, sin dependencias.
(() => {
  'use strict';

  // Año dinámico en el pie.
  const y = document.getElementById('anio');
  if (y) y.textContent = String(new Date().getFullYear());

  // Detección simple de SO para adaptar la extensión del instalador.
  const ua = navigator.userAgent;
  const os = /Mac/i.test(ua) ? { label: 'macOS', ext: '.dmg' }
    : /Linux/i.test(ua) ? { label: 'Linux', ext: '.AppImage' }
    : { label: 'Windows 10/11', ext: '.exe' };

  const ext = document.getElementById('ext');
  const osLabel = document.getElementById('os-label');
  if (ext) ext.textContent = os.ext;
  if (osLabel) osLabel.textContent = os.label;
})();
