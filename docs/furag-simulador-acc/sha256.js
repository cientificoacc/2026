/* SHA-256 en JavaScript puro.
   Se usa esta implementación y no crypto.subtle porque el aplicativo debe funcionar
   también abierto desde el disco (file://), donde algunos navegadores no exponen la
   Web Crypto API. Solo interviene en el modo local; en modo remoto la verificación
   de la contraseña ocurre en el servidor. */
window.sha256FURAG = (function () {
  'use strict';
  var K = [], H0 = [], primos = [], n = 2;
  function raiz(x, e) { return Math.floor((Math.pow(x, e) % 1) * Math.pow(2, 32)); }
  while (primos.length < 64) {
    var esP = true;
    for (var d = 2; d * d <= n; d++) if (n % d === 0) { esP = false; break; }
    if (esP) primos.push(n);
    n++;
  }
  for (var i = 0; i < 64; i++) K[i] = raiz(primos[i], 1 / 3) | 0;
  for (var j = 0; j < 8; j++) H0[j] = raiz(primos[j], 1 / 2) | 0;
  function rotr(x, k) { return (x >>> k) | (x << (32 - k)); }
  function utf8(s) {
    var b = [], c;
    for (var i = 0; i < s.length; i++) {
      c = s.charCodeAt(i);
      if (c < 0x80) b.push(c);
      else if (c < 0x800) b.push(0xc0 | (c >> 6), 0x80 | (c & 63));
      else if (c < 0xd800 || c >= 0xe000) b.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
      else {
        c = 0x10000 + (((c & 0x3ff) << 10) | (s.charCodeAt(++i) & 0x3ff));
        b.push(0xf0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
      }
    }
    return b;
  }
  return function (msg) {
    var b = utf8(msg), l = b.length * 8;
    b.push(0x80);
    while (b.length % 64 !== 56) b.push(0);
    b.push(0, 0, 0, 0, (l >>> 24) & 255, (l >>> 16) & 255, (l >>> 8) & 255, l & 255);
    var H = H0.slice(), w = new Array(64);
    for (var i = 0; i < b.length; i += 64) {
      for (var t = 0; t < 16; t++)
        w[t] = (b[i + t * 4] << 24) | (b[i + t * 4 + 1] << 16) | (b[i + t * 4 + 2] << 8) | b[i + t * 4 + 3];
      for (t = 16; t < 64; t++) {
        var s0 = rotr(w[t - 15], 7) ^ rotr(w[t - 15], 18) ^ (w[t - 15] >>> 3);
        var s1 = rotr(w[t - 2], 17) ^ rotr(w[t - 2], 19) ^ (w[t - 2] >>> 10);
        w[t] = (w[t - 16] + s0 + w[t - 7] + s1) | 0;
      }
      var a = H[0], bb = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
      for (t = 0; t < 64; t++) {
        var S1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
        var ch = (e & f) ^ (~e & g);
        var t1 = (h + S1 + ch + K[t] + w[t]) | 0;
        var S0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
        var mj = (a & bb) ^ (a & c) ^ (bb & c);
        var t2 = (S0 + mj) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = bb; bb = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0] + a) | 0; H[1] = (H[1] + bb) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
      H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
    }
    return H.map(function (x) { return ('00000000' + (x >>> 0).toString(16)).slice(-8); }).join('');
  };
})();
