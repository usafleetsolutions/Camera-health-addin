/*
 * USAFS Camera Health Add-In
 *
 * Displays the current database's own latest weekly camera health report inside
 * MyGeotab. Does NOT compute anything itself -- it asks a small lookup proxy
 * (apps/camera-addin/apps-script/Code.gs) "what's the report for THIS database?", keyed
 * by the database name MyGeotab's own injected `api` object reports for the logged-in
 * session. The proxy sits in front of a private index spreadsheet (see
 * scripts/publish_camera_addin_index.py) and answers one row at a time -- it never
 * returns every customer's report link in one response, which an earlier version of
 * this add-in did (a public sheet fetched wholesale client-side). See
 * plans/2026-07-08-geotab-camera-health-addin.md, Design Decision 5.
 *
 * LOOKUP_PROXY_URL is a placeholder until the Apps Script proxy has been deployed --
 * see apps/camera-addin/apps-script/README.md. Fill it in before deploying this file.
 */

var LOOKUP_PROXY_URL = "https://script.google.com/macros/s/AKfycbyQSmpTatHxXQHzv7MD_l_PE4T4p6oFtRi2r8iaxtGL4bu7fqXDHYA5anlDn_g-zWbo/exec";

// Confirmed live (2026-07-09), reused from an earlier logo-sizing spike in this same
// repo (logo-test.html/.js) -- USA Fleet Solutions' own marketing site, not something
// we host ourselves.
var LOGO_URL = "https://www.usafleetsolutions.com/wp-content/themes/usa-fleet/images/logo.png";

geotab.addin.usafsCameraHealth = function () {
  "use strict";

  var elRoot = null;
  var contentRoot = null;

  function el(tag, attrs, children) {
    var node = document.createElement(tag);
    attrs = attrs || {};
    for (var key in attrs) {
      if (key === "text") {
        node.textContent = attrs[key];
      } else {
        node.setAttribute(key, attrs[key]);
      }
    }
    (children || []).forEach(function (child) {
      node.appendChild(child);
    });
    return node;
  }

  // Every style below is INLINE (a "style" attribute), not a CSS class. MyGeotab
  // appears to mirror/clone this page's rendered DOM nodes into its own page rather
  // than displaying our iframe directly -- confirmed during pilot testing on
  // hunt_sons (2026-07-08): repeated changes to the <style> block in index.html had
  // zero visual effect even after confirming via curl the updated CSS was genuinely
  // live. A cloned node carries its own inline style with it; it does not carry a
  // separate stylesheet's rules along. min-height is used instead of percentage-based
  // height throughout, since percentage height depends on an ancestor's computed
  // height, and we have no visibility into (or control over) whatever MyGeotab's own
  // wrapper element around the mirrored content actually is.

  // Shown once, above whatever the current state is (loading / error / report) --
  // per Paul (2026-07-09): needs to be obviously "USA Fleet Solutions" branded (this
  // is a differentiator vs. other Geotab resellers, most of whom offer nothing like
  // it) but with a short name and a compact logo, since there's limited real estate
  // above the actual report content.
  function renderBrandHeader() {
    return el("div", {
      style: "display:flex;align-items:center;gap:12px;padding:2px 0 12px 0;margin-bottom:12px;border-bottom:1px solid #eee;",
    }, [
      el("img", {
        src: LOGO_URL,
        alt: "USA Fleet Solutions",
        style: "display:block;width:110px;height:auto;flex-shrink:0;",
      }),
      el("div", {}, [
        el("div", {
          style: "font-size:15px;font-weight:700;color:#003087;line-height:1.2;",
          text: "USA Fleet Camera Health",
        }),
        el("div", {
          style: "font-size:11px;color:#888;line-height:1.4;margin-top:3px;",
          text: "A USA Fleet Solutions exclusive add-in for our customers. Issues? Contact support@usafleetsolutions.com.",
        }),
      ]),
    ]);
  }

  function renderMessage(text) {
    contentRoot.innerHTML = "";
    contentRoot.appendChild(el("div", {
      style: "padding:24px 8px;font-size:14px;color:#555;line-height:1.5;",
    }, [
      document.createTextNode(text),
    ]));
  }

  // hasHtml=true: embed the real report, served live by the lookup proxy (reads a
  // private Drive file server-side -- see Code.gs). hasHtml=false: fall back to the
  // Drive PDF preview (older customers whose HTML hasn't been privately uploaded yet).
  function renderReport(databaseName, reportUrl, generatedAt, hasHtml) {
    contentRoot.innerHTML = "";

    var meta = el("div", {
      style: "font-size:12px;color:#888;margin-bottom:6px;",
    }, [
      document.createTextNode(generatedAt ? "Report generated: " + generatedAt : ""),
    ]);
    contentRoot.appendChild(meta);

    var actionsRow = el("div", {
      style: "display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:10px;",
    });

    if (reportUrl) {
      actionsRow.appendChild(el("a", {
        style: "font-size:13px;color:#2980b9;text-decoration:none;",
        href: reportUrl,
        target: "_blank",
        rel: "noopener",
        text: "Open as PDF →",
      }));
    }

    // On-demand refresh (added 2026-07-09): queues a one-off regeneration of just this
    // database's report on Paul's machine -- see Code.gs's action=refresh handling and
    // scripts/process_refresh_queue.py. Rate-limited to once per 24h per database
    // server-side, so this button can't be spammed; the response message (queued /
    // already-requested-today / temporarily-unavailable) is shown as-is, not
    // reinterpreted here, so the messaging only needs to be right in one place.
    var refreshStatus = el("span", { style: "font-size:12px;color:#888;" });
    var refreshBtn = el("button", {
      type: "button",
      style: "font-size:12px;color:#2980b9;background:#eef5fb;border:1px solid #cfe2f3;border-radius:4px;padding:4px 10px;cursor:pointer;",
      text: "Refresh Report",
    });
    refreshBtn.addEventListener("click", function () {
      requestRefresh(databaseName, refreshBtn, refreshStatus);
    });
    actionsRow.appendChild(refreshBtn);
    actionsRow.appendChild(refreshStatus);

    contentRoot.appendChild(actionsRow);

    var frameSrc = hasHtml
      ? LOOKUP_PROXY_URL + "?db=" + encodeURIComponent(databaseName) + "&view=report"
      : toDrivePreviewUrl(reportUrl);

    contentRoot.appendChild(el("iframe", {
      style: "display:block;width:100%;min-height:600px;flex:1 1 auto;border:1px solid #e0e0e0;border-radius:6px;",
      src: frameSrc,
      title: "Camera Health Report",
    }));
  }

  function requestRefresh(databaseName, btn, statusEl) {
    btn.disabled = true;
    var originalLabel = btn.textContent;
    btn.textContent = "Requesting…";
    statusEl.textContent = "";

    var url = LOOKUP_PROXY_URL + "?db=" + encodeURIComponent(databaseName) + "&action=refresh";
    fetch(url)
      .then(function (resp) {
        if (!resp.ok) throw new Error("request failed: " + resp.status);
        return resp.json();
      })
      .then(function (result) {
        statusEl.textContent = result.message || "Refresh requested.";
        statusEl.style.color = (result.status === "unavailable") ? "#e74c3c" : "#555";
      })
      .catch(function (err) {
        statusEl.textContent = "Couldn't request a refresh right now (" + err.message + "). Try again shortly.";
        statusEl.style.color = "#e74c3c";
      })
      .finally(function () {
        btn.disabled = false;
        btn.textContent = originalLabel;
      });
  }

  // Google Drive "webViewLink" URLs look like .../file/d/{id}/view?usp=... -- the
  // /preview variant is the one that's embeddable in an iframe. Only used as a fallback
  // now that the proxy can serve the real HTML report directly (see renderReport).
  function toDrivePreviewUrl(url) {
    return url.replace(/\/view(\?.*)?$/, "/preview");
  }

  // ── Dual-vendor support (added 2026-07-14) ──────────────────────────────────
  // A small number of mid-migration customers (Western Turf, TAC Transportation)
  // genuinely run both Geotab-side cameras (GO Focus/Surfsight) AND VisionTrack at
  // once -- per Paul, these should show as two lines/two reports, not one merged
  // view. The lookup proxy always answers by this database's OWN key (the only
  // thing api.getSession() can tell us -- VisionTrack orgs aren't part of MyGeotab
  // at all), but a dual-vendor row's response carries a second report (has_vt=true,
  // vt_report_url/vt_generated_at/has_vt_html/vt_monitored/vt_key) riding alongside
  // the primary one. See publish_camera_addin_index.py's merge logic.
  //
  // renderReport() above stays untouched and is still used for the ~99 single-vendor
  // customers -- this is a separate path so that population has zero behavior change.

  function tabButtonStyle(active) {
    return "font-size:13px;font-weight:" + (active ? "700" : "400") +
      ";color:" + (active ? "#003087" : "#888") +
      ";background:none;border:none;border-bottom:2px solid " + (active ? "#003087" : "transparent") +
      ";padding:6px 4px;cursor:pointer;";
  }

  // Builds one tab's content: a "not ready yet" message if that system's report
  // isn't monitored/generated yet, otherwise the same meta/actions/iframe structure
  // renderReport() uses for the single-report case. `databaseName` is always the
  // Geotab database (that's the only valid ?db= value the proxy accepts for
  // &view=report); `refreshKey` is what gets sent for &action=refresh, which DOES
  // differ per system (the Geotab database for one tab, "visiontrack-<prefix>" for
  // the other) since process_refresh_queue.py dispatches on that key's prefix.
  function buildSystemPanel(databaseName, refreshKey, which, info) {
    var panel = el("div", {});

    if (!info.monitored || (!info.reportUrl && !info.hasHtml)) {
      panel.appendChild(el("div", {
        style: "padding:24px 8px;font-size:14px;color:#555;line-height:1.5;",
      }, [
        document.createTextNode(
          which === "vt"
            ? "VisionTrack camera health hasn't been generated for this account yet. Check back after the next weekly run."
            : "This database is on the camera health list, but a report hasn't been generated yet. Check back after the next weekly run."
        ),
      ]));
      return panel;
    }

    var meta = el("div", {
      style: "font-size:12px;color:#888;margin-bottom:6px;",
    }, [
      document.createTextNode(info.generatedAt ? "Report generated: " + info.generatedAt : ""),
    ]);
    panel.appendChild(meta);

    var actionsRow = el("div", {
      style: "display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:10px;",
    });

    if (info.reportUrl) {
      actionsRow.appendChild(el("a", {
        style: "font-size:13px;color:#2980b9;text-decoration:none;",
        href: info.reportUrl,
        target: "_blank",
        rel: "noopener",
        text: "Open as PDF →",
      }));
    }

    var refreshStatus = el("span", { style: "font-size:12px;color:#888;" });
    var refreshBtn = el("button", {
      type: "button",
      style: "font-size:12px;color:#2980b9;background:#eef5fb;border:1px solid #cfe2f3;border-radius:4px;padding:4px 10px;cursor:pointer;",
      text: "Refresh Report",
    });
    refreshBtn.addEventListener("click", function () {
      requestRefresh(refreshKey, refreshBtn, refreshStatus);
    });
    actionsRow.appendChild(refreshBtn);
    actionsRow.appendChild(refreshStatus);
    panel.appendChild(actionsRow);

    var frameSrc = info.hasHtml
      ? LOOKUP_PROXY_URL + "?db=" + encodeURIComponent(databaseName) + "&view=report&which=" + which
      : toDrivePreviewUrl(info.reportUrl);

    panel.appendChild(el("iframe", {
      style: "display:block;width:100%;min-height:600px;flex:1 1 auto;border:1px solid #e0e0e0;border-radius:6px;",
      src: frameSrc,
      title: which === "vt" ? "VisionTrack Camera Health Report" : "Camera Health Report",
    }));

    return panel;
  }

  function renderDualReport(databaseName, vtRefreshKey, geotabInfo, vtInfo) {
    contentRoot.innerHTML = "";

    var tabBar = el("div", {
      style: "display:flex;gap:16px;margin-bottom:10px;border-bottom:1px solid #eee;",
    });
    var geotabTabBtn = el("button", { type: "button", style: tabButtonStyle(false), text: "Geotab Cameras" });
    var vtTabBtn = el("button", { type: "button", style: tabButtonStyle(false), text: "VisionTrack Cameras" });
    tabBar.appendChild(geotabTabBtn);
    tabBar.appendChild(vtTabBtn);

    var geotabPanel = buildSystemPanel(databaseName, databaseName, "geotab", geotabInfo);
    var vtPanel = buildSystemPanel(databaseName, vtRefreshKey, "vt", vtInfo);

    function activate(which) {
      var isGeotab = which === "geotab";
      geotabPanel.style.display = isGeotab ? "" : "none";
      vtPanel.style.display = isGeotab ? "none" : "";
      geotabTabBtn.setAttribute("style", tabButtonStyle(isGeotab));
      vtTabBtn.setAttribute("style", tabButtonStyle(!isGeotab));
    }

    geotabTabBtn.addEventListener("click", function () { activate("geotab"); });
    vtTabBtn.addEventListener("click", function () { activate("vt"); });

    contentRoot.appendChild(tabBar);
    contentRoot.appendChild(geotabPanel);
    contentRoot.appendChild(vtPanel);

    // Default to whichever side actually has a ready report, in case one system's
    // weekly run hasn't produced anything yet for this brand-new dual customer.
    activate(geotabInfo.monitored ? "geotab" : (vtInfo.monitored ? "vt" : "geotab"));
  }

  // notifyReady() is MyGeotab's "I'm initialized" signal (the callback passed into
  // initialize()). Earlier versions called it synchronously, immediately, before any
  // real content existed -- pilot testing on hunt_sons (2026-07-08) showed MyGeotab
  // renders the add-in in a small, fixed, scrollable box that never grows afterward,
  // strongly suggesting MyGeotab sizes/finalizes its container around whatever content
  // exists at the moment this callback fires, and doesn't re-measure later. Every
  // render path below (success, not-found, not-monitored, error) now calls this AFTER
  // painting real content, and guards against double-calling since only one path runs.
  var readyCalled = false;
  function notifyReady(callback) {
    if (readyCalled) return;
    readyCalled = true;
    // One paint/layout tick so the DOM mutation above is actually reflected before
    // MyGeotab measures -- requestAnimationFrame guarantees a layout has happened,
    // a plain synchronous call right after innerHTML changes might not.
    requestAnimationFrame(function () {
      callback();
    });
  }

  function loadReport(databaseName, callback) {
    if (LOOKUP_PROXY_URL.indexOf("REPLACE_WITH") === 0) {
      renderMessage("Camera Health Add-In isn't fully configured yet -- ask your account manager to check back soon.");
      notifyReady(callback);
      return;
    }

    var url = LOOKUP_PROXY_URL + "?db=" + encodeURIComponent(databaseName);

    fetch(url)
      .then(function (resp) {
        if (!resp.ok) throw new Error("lookup failed: " + resp.status);
        return resp.json();
      })
      .then(function (match) {
        if (!match.found) {
          renderMessage("Camera health monitoring isn't set up for this database yet. Contact your account manager if you'd like to be added.");
          return;
        }
        if (match.has_vt) {
          // Dual-vendor customer -- show both systems as separate tabs instead of
          // picking one. Each tab handles its own "not ready yet" state internally
          // (buildSystemPanel), so neither system being ready shouldn't block the
          // other from displaying.
          renderDualReport(
            databaseName,
            match.vt_key || databaseName,
            { reportUrl: match.report_url, generatedAt: match.generated_at, hasHtml: match.has_html, monitored: !!match.monitored },
            { reportUrl: match.vt_report_url, generatedAt: match.vt_generated_at, hasHtml: match.has_vt_html, monitored: !!match.vt_monitored }
          );
          return;
        }
        if (!match.monitored || (!match.report_url && !match.has_html)) {
          renderMessage("This database is on the camera health list, but a report hasn't been generated yet. Check back after the next weekly run.");
          return;
        }
        renderReport(databaseName, match.report_url, match.generated_at, match.has_html);
      })
      .catch(function (err) {
        renderMessage("Couldn't load the camera health report right now (" + err.message + "). Try again shortly, or contact your account manager.");
      })
      .finally(function () {
        notifyReady(callback);
      });
  }

  return {
    initialize: function (api, state, callback) {
      elRoot = document.getElementById("usafsCameraHealthRoot");
      elRoot.innerHTML = "";
      elRoot.appendChild(renderBrandHeader());
      contentRoot = el("div", {});
      elRoot.appendChild(contentRoot);

      renderMessage("Loading your camera health report…");

      // api.getSession() inside a real MyGeotab add-in only supports the single-
      // callback form -- passing a second (error) callback throws a synchronous
      // "MethodNotSupported" before anything else runs, which is what happened here
      // during pilot testing on hunt_sons (2026-07-08). Wrap in try/catch instead.
      try {
        api.getSession(function (session) {
          loadReport(session.database, callback);
        });
      } catch (err) {
        renderMessage("Couldn't identify this database's session (" + err.message + "). Try reloading the page.");
        notifyReady(callback);
      }
    },

    focus: function () {},

    blur: function () {},
  };
};
