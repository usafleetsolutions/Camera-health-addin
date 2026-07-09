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

geotab.addin.usafsCameraHealth = function () {
  "use strict";

  var elRoot = null;

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

  function renderMessage(text) {
    elRoot.innerHTML = "";
    elRoot.appendChild(el("div", { class: "usafs-camera-message" }, [
      document.createTextNode(text),
    ]));
  }

  // hasHtml=true: embed the real report, served live by the lookup proxy (reads a
  // private Drive file server-side -- see Code.gs). hasHtml=false: fall back to the
  // Drive PDF preview (older customers whose HTML hasn't been privately uploaded yet).
  function renderReport(databaseName, reportUrl, generatedAt, hasHtml) {
    elRoot.innerHTML = "";

    var meta = el("div", { class: "usafs-camera-meta" }, [
      document.createTextNode(generatedAt ? "Report generated: " + generatedAt : ""),
    ]);
    elRoot.appendChild(meta);

    if (reportUrl) {
      elRoot.appendChild(el("a", {
        class: "usafs-camera-open-link",
        href: reportUrl,
        target: "_blank",
        rel: "noopener",
        text: "Open as PDF →",
      }));
    }

    var frameSrc = hasHtml
      ? LOOKUP_PROXY_URL + "?db=" + encodeURIComponent(databaseName) + "&view=report"
      : toDrivePreviewUrl(reportUrl);

    elRoot.appendChild(el("iframe", {
      class: "usafs-camera-frame",
      src: frameSrc,
      title: "Camera Health Report",
    }));
  }

  // Google Drive "webViewLink" URLs look like .../file/d/{id}/view?usp=... -- the
  // /preview variant is the one that's embeddable in an iframe. Only used as a fallback
  // now that the proxy can serve the real HTML report directly (see renderReport).
  function toDrivePreviewUrl(url) {
    return url.replace(/\/view(\?.*)?$/, "/preview");
  }

  function loadReport(databaseName) {
    if (LOOKUP_PROXY_URL.indexOf("REPLACE_WITH") === 0) {
      renderMessage("Camera Health Add-In isn't fully configured yet -- ask your account manager to check back soon.");
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
        if (!match.monitored || (!match.report_url && !match.has_html)) {
          renderMessage("This database is on the camera health list, but a report hasn't been generated yet. Check back after the next weekly run.");
          return;
        }
        renderReport(databaseName, match.report_url, match.generated_at, match.has_html);
      })
      .catch(function (err) {
        renderMessage("Couldn't load the camera health report right now (" + err.message + "). Try again shortly, or contact your account manager.");
      });
  }

  return {
    initialize: function (api, state, callback) {
      elRoot = document.getElementById("usafsCameraHealthRoot");
      renderMessage("Loading your camera health report…");

      // api.getSession() inside a real MyGeotab add-in only supports the single-
      // callback form -- passing a second (error) callback throws a synchronous
      // "MethodNotSupported" before anything else runs, which is what happened here
      // during pilot testing on hunt_sons (2026-07-08). Wrap in try/catch instead.
      try {
        api.getSession(function (session) {
          loadReport(session.database);
        });
      } catch (err) {
        renderMessage("Couldn't identify this database's session (" + err.message + "). Try reloading the page.");
      }

      callback();
    },

    focus: function () {},

    blur: function () {},
  };
};
