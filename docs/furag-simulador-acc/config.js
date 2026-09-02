/* ─────────────────────────────────────────────────────────────────────────────
   Configuración del Simulador FURAG · Agencia Catastral de Cundinamarca

   api  → URL del Web App de Google Apps Script que hace de backend.
          · Vacío  = modo local: cada navegador guarda su propio avance. Sirve para
            probar el aplicativo con doble clic, sin instalar nada.
          · Con URL = modo compartido: todas las dependencias escriben en la misma
            hoja de cálculo y ven el avance de las demás.

   Cómo obtener esa URL: backend/INSTALACION.md
   Pegue aquí la dirección que termina en /exec, entre las comillas.
   ───────────────────────────────────────────────────────────────────────────── */
window.FURAG_CONFIG = {
  api: 'https://script.google.com/macros/s/AKfycbzZpIKEpNRMeBtjrohluiM5QB8FmqzBFZRZw1mgYEhol6tnR4QzftmX-jYYELy_J0lH/exec',

  /* Cada cuántos segundos consulta el aplicativo si hubo cambios de otras
     dependencias. Solo aplica en modo compartido. */
  sondeoSegundos: 60
};
