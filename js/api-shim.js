/**
 * api-shim.js - Adapter agar kode lama (google.script.run.namaFungsi(...)) tetap jalan
 * di luar Apps Script. Setiap pemanggilan diterjemahkan menjadi POST JSON ke Web App.
 *
 *  - Token login disimpan di MEMORI saja (bukan localStorage) -> refresh halaman = login ulang,
 *    sama seperti perilaku aplikasi sebelumnya, dan lebih kecil risikonya bila ada XSS.
 *  - Content-Type "text/plain" sengaja dipakai agar browser tidak mengirim preflight CORS
 *    (Apps Script tidak mendukung request OPTIONS).
 */
(function () {
  'use strict';

  var cfg = window.APP_CONFIG || {};
  var token = null;

  // Batas waktu per fungsi (ms). Peta bisa lama karena geocoding alamat.
  var DEFAULT_TIMEOUT = 90 * 1000;
  var TIMEOUTS = { getRelokasiMapPoints: 340 * 1000 };

  var API = {
    clearSession: function () { token = null; },
    hasSession: function () { return !!token; },
    // Dipanggil saat server menolak token (kedaluwarsa/tidak valid). Boleh ditimpa.
    onSessionExpired: function (message) {
      alert((message || 'Sesi Anda berakhir.') + '\nHalaman akan dimuat ulang.');
      window.location.reload();
    }
  };
  window.APP_API = API;

  function call(fnName, args) {
    if (!cfg.API_URL || cfg.API_URL.indexOf('GANTI_DENGAN') !== -1) {
      return Promise.reject(new Error('API_URL belum diatur. Edit file js/config.js.'));
    }
    var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = controller ? setTimeout(function () { controller.abort(); }, TIMEOUTS[fnName] || DEFAULT_TIMEOUT) : null;

    return fetch(cfg.API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fnName, args: args, token: token }),
      redirect: 'follow',
      signal: controller ? controller.signal : undefined
    }).then(function (res) {
      return res.text();
    }).then(function (text) {
      var json;
      try { json = JSON.parse(text); }
      catch (e) { throw new Error('Respons server tidak valid. Periksa API_URL dan pastikan deployment Apps Script sudah diperbarui.'); }

      if (!json.ok) {
        if (json.code === 'AUTH' && fnName !== 'login') {
          token = null;
          API.onSessionExpired(json.error);
        }
        throw new Error(json.error || 'Terjadi kesalahan pada server.');
      }
      if (fnName === 'login' && json.data && json.data.success && json.data.token) {
        token = json.data.token;
        delete json.data.token; // token tidak perlu terlihat oleh kode aplikasi lain
      }
      return json.data;
    }).catch(function (err) {
      if (err && err.name === 'AbortError') throw new Error('Server terlalu lama merespons. Coba lagi.');
      if (err instanceof TypeError) throw new Error('Tidak dapat terhubung ke server. Periksa koneksi internet.');
      throw err;
    }).then(function (data) {
      if (timer) clearTimeout(timer);
      return data;
    }, function (err) {
      if (timer) clearTimeout(timer);
      throw err;
    });
  }

  // Meniru: google.script.run.withSuccessHandler(fn).withFailureHandler(fn).namaFungsi(...)
  function createRunner(onSuccess, onFailure) {
    return new Proxy({}, {
      get: function (_target, name) {
        if (typeof name !== 'string') return undefined;
        if (name === 'withSuccessHandler') return function (fn) { return createRunner(fn, onFailure); };
        if (name === 'withFailureHandler') return function (fn) { return createRunner(onSuccess, fn); };
        if (name === 'withUserObject') return function () { return createRunner(onSuccess, onFailure); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          call(name, args).then(function (data) {
            if (onSuccess) onSuccess(data);
          }, function (err) {
            if (onFailure) onFailure(err);
            else if (window.console) console.error('[API] ' + name + ':', err);
          });
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = { run: createRunner(null, null) };
})();
