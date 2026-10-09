/* Back-arrow history for tools that swap views in place.

   Every tool's top bar has a back arrow that is just history.back(). A tool that switches tabs, wizard steps or
   screens without a page load never created a history step, so the arrow skipped straight past it to the page
   before the tool. NavHist makes each view one history step: the arrow (and the browser's back button) then walk
   back through the views you opened, and only leave the tool once there is nothing left to step back to.

   Usage, once the tool's views exist:
       NavHist.init({ snapshot: () => ({ ...the current view... }), restore: s => { ...show that view... } });
   and after any view change (inside the function that makes the change):
       NavHist.record();
   snapshot() must return a small JSON-friendly object. restore() is only called while going back/forward, and
   NavHist.record() does nothing during it, so restoring never adds steps of its own. */
(function () {
  var cfg = null, last = '', restoring = false, quietUntil = 0;

  function snap() {
    try { return cfg.snapshot(); } catch (e) { return null; }
  }
  function ser(s) { return JSON.stringify(s); }

  window.NavHist = {
    get restoring() { return restoring; },
    init: function (c) {
      cfg = c;
      var s = snap();
      var saved = history.state && history.state.navHist;
      if (c.restoreOnLoad && saved && ser(saved) !== ser(s)) {
        // Coming back to this page (or refreshing it): reopen the view this history entry was on.
        restoring = true;
        try { c.restore(saved); } catch (err) { console.error('NavHist restore failed', err); }
        finally { restoring = false; s = snap() || s; }
      }
      last = ser(s);
      try { history.replaceState(Object.assign({}, history.state, { navHist: s }), ''); } catch (e) { /* sandboxed */ }
      window.addEventListener('popstate', function (e) {
        var target = e.state && e.state.navHist;
        if (!target || !cfg) return;
        restoring = true;
        try { cfg.restore(target); }
        catch (err) { console.error('NavHist restore failed', err); }
        // A restore can finish drawing a moment later (fetches, animations); ignore those redraws so they don't
        // count as new steps.
        finally { restoring = false; last = ser(target); quietUntil = Date.now() + 800; }
      });
    },
    record: function () {
      if (!cfg || restoring || Date.now() < quietUntil) return;
      var s = snap();
      if (!s) return;
      var k = ser(s);
      if (k === last) return;
      last = k;
      try { history.pushState(Object.assign({}, history.state, { navHist: s }), ''); } catch (e) { /* sandboxed */ }
    },
  };
})();
