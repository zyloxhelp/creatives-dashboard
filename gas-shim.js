// Lets the form run on GitHub Pages. It gives the page the same google.script.run calls it uses
// inside Apps Script, but sends them to the deployed web app URL (config.js) with fetch.
(function () {
  var NAMES = ['getConfig', 'submitRequest', 'getTracker', 'updateStatus'];
  function call(fn, args) {
    var url = String(window.PG_API_URL || '').trim();
    if (!url) return Promise.reject(new Error('Add your Apps Script web app URL in config.js (PG_API_URL).'));
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // plain text keeps it a simple request, no CORS preflight
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    }).then(function (r) { return r.text(); }).then(function (t) {
      var j;
      try { j = JSON.parse(t); } catch (e) { throw new Error('The web app did not return data. Check it is deployed with access set to Anyone, and that the URL ends in /exec.'); }
      if (!j.ok) throw new Error(j.error || 'Request failed');
      return j.data;
    });
  }
  function runner(ok, fail) {
    var r = {
      withSuccessHandler: function (f) { return runner(f, fail); },
      withFailureHandler: function (f) { return runner(ok, f); }
    };
    NAMES.forEach(function (n) {
      r[n] = function () {
        call(n, Array.prototype.slice.call(arguments)).then(function (v) { if (ok) ok(v); }, function (e) { if (fail) fail(e); });
      };
    });
    return r;
  }
  window.google = { script: { run: runner(null, null) } };
})();
