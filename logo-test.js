geotab.addin.usafsLogoTest = function () {
  "use strict";

  var elRoot = null;

  // MyGeotab clones DOM nodes -- only inline style attributes work, not CSS classes.
  // Call callback AFTER painting so Geotab sizes the container around real content.
  return {
    initialize: function (api, state, callback) {
      elRoot = document.getElementById("usafsLogoTestRoot");

      // Paint the logo immediately -- no async needed for this test
      elRoot.innerHTML = "";
      elRoot.appendChild((function () {
        var img = document.createElement("img");
        img.setAttribute("src", "https://www.usafleetsolutions.com/wp-content/themes/usa-fleet/images/logo.png");
        img.setAttribute("alt", "USA Fleet Solutions");
        img.setAttribute("style", "display:block;max-width:100%;width:480px;height:auto;");
        img.onerror = function () {
          // Fallback: text logo if image URL is wrong
          var fallback = document.createElement("div");
          fallback.setAttribute("style", "font-size:28px;font-weight:700;color:#003087;letter-spacing:1px;text-align:center;");
          fallback.textContent = "USA Fleet Solutions";
          elRoot.innerHTML = "";
          elRoot.appendChild(fallback);
        };
        return img;
      })());

      // One paint tick before signaling ready so Geotab measures real content
      requestAnimationFrame(function () {
        callback();
      });
    },

    focus: function () {},
    blur: function () {},
  };
};
