/* =========================================================================
   sheet-lite docs — minimal vanilla JS, no dependencies, no network.
   Handles: theme toggle, copy-to-clipboard buttons, mobile nav drawer,
   active-link highlighting, and accessible scroll behaviour.
   ========================================================================= */
(function () {
  "use strict";

  var THEME_KEY = "sheet-lite:theme";
  var root = document.documentElement;

  /* ------------------------------------------------------------ theme */
  function applyTheme(theme, persist) {
    root.setAttribute("data-theme", theme);
    if (persist) {
      try {
        localStorage.setItem(THEME_KEY, theme);
      } catch (e) {
        /* storage may be unavailable */
      }
    }
    var toggle = document.getElementById("theme-toggle");
    if (toggle) {
      toggle.setAttribute(
        "aria-label",
        theme === "dark" ? "Switch to light theme" : "Switch to dark theme"
      );
    }
  }

  function setTheme(theme) {
    applyTheme(theme, true);
  }

  function currentTheme() {
    return root.getAttribute("data-theme") === "dark" ? "dark" : "light";
  }

  var themeToggle = document.getElementById("theme-toggle");
  if (themeToggle) {
    themeToggle.addEventListener("click", function () {
      setTheme(currentTheme() === "dark" ? "light" : "dark");
    });
  }

  // Follow the OS preference when the user hasn't chosen explicitly.
  if (window.matchMedia) {
    var mq = window.matchMedia("(prefers-color-scheme: dark)");
    var onSchemeChange = function (event) {
      var stored = null;
      try {
        stored = localStorage.getItem(THEME_KEY);
      } catch (e) {
        /* ignore */
      }
      if (!stored) applyTheme(event.matches ? "dark" : "light", false);
    };
    if (mq.addEventListener) mq.addEventListener("change", onSchemeChange);
    else if (mq.addListener) mq.addListener(onSchemeChange);
  }

  /* -------------------------------------------------- copy buttons */
  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    // Fallback for file:// and older browsers.
    return new Promise(function (resolve, reject) {
      try {
        var area = document.createElement("textarea");
        area.value = text;
        area.setAttribute("readonly", "");
        area.style.position = "fixed";
        area.style.top = "-1000px";
        document.body.appendChild(area);
        area.select();
        var ok = document.execCommand("copy");
        document.body.removeChild(area);
        ok ? resolve() : reject(new Error("copy failed"));
      } catch (err) {
        reject(err);
      }
    });
  }

  function enhanceCodeBlocks() {
    var blocks = document.querySelectorAll(".code-block");
    Array.prototype.forEach.call(blocks, function (block) {
      var pre = block.querySelector("pre");
      if (!pre || block.querySelector(".copy-btn")) return;

      var header = block.querySelector(".code-block__header");
      if (!header) {
        header = document.createElement("div");
        header.className = "code-block__header";
        block.insertBefore(header, pre);
      }

      var btn = document.createElement("button");
      btn.type = "button";
      btn.className = "copy-btn";
      btn.textContent = "Copy";
      btn.setAttribute("aria-label", "Copy code to clipboard");

      btn.addEventListener("click", function () {
        var code = pre.querySelector("code");
        var text = (code ? code.textContent : pre.textContent) || "";
        copyText(text).then(
          function () {
            btn.textContent = "Copied";
            btn.classList.add("is-copied");
            window.setTimeout(function () {
              btn.textContent = "Copy";
              btn.classList.remove("is-copied");
            }, 1600);
          },
          function () {
            btn.textContent = "Press Ctrl+C";
            window.setTimeout(function () {
              btn.textContent = "Copy";
            }, 1600);
          }
        );
      });

      header.appendChild(btn);
    });
  }

  /* ------------------------------------------------ mobile nav drawer */
  var sidebar = document.getElementById("sidebar");
  var menuToggle = document.getElementById("menu-toggle");
  var scrim = document.createElement("div");
  scrim.className = "scrim";
  document.body.appendChild(scrim);

  function setDrawer(open) {
    if (!sidebar || !menuToggle) return;
    sidebar.classList.toggle("is-open", open);
    scrim.classList.toggle("is-visible", open);
    menuToggle.setAttribute("aria-expanded", open ? "true" : "false");
  }

  if (menuToggle) {
    menuToggle.addEventListener("click", function () {
      setDrawer(!sidebar.classList.contains("is-open"));
    });
  }
  scrim.addEventListener("click", function () {
    setDrawer(false);
  });
  document.addEventListener("keydown", function (event) {
    if (event.key === "Escape") setDrawer(false);
  });

  // Close the drawer after choosing a link on small screens.
  Array.prototype.forEach.call(
    document.querySelectorAll(".nav__link"),
    function (link) {
      link.addEventListener("click", function () {
        if (window.matchMedia && window.matchMedia("(max-width: 820px)").matches) {
          setDrawer(false);
        }
      });
    }
  );

  /* ------------------------------------------------ active nav link */
  function enhanceActiveLinks() {
    var links = Array.prototype.slice.call(
      document.querySelectorAll(".nav__link")
    );
    if (!links.length || !("IntersectionObserver" in window)) return;

    var byId = {};
    var sections = [];
    links.forEach(function (link) {
      var href = link.getAttribute("href") || "";
      if (href.charAt(0) !== "#") return;
      var id = href.slice(1);
      var el = document.getElementById(id);
      if (!el) return;
      byId[id] = link;
      sections.push(el);
    });

    var visible = {};
    var observer = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          visible[entry.target.id] = entry.isIntersecting
            ? entry.intersectionRatio
            : 0;
        });

        var bestId = null;
        var bestRatio = 0;
        Object.keys(visible).forEach(function (id) {
          if (visible[id] > bestRatio) {
            bestRatio = visible[id];
            bestId = id;
          }
        });

        if (bestId) {
          links.forEach(function (link) {
            link.classList.remove("is-active");
          });
          if (byId[bestId]) byId[bestId].classList.add("is-active");
        }
      },
      {
        rootMargin: "-" + 72 + "px 0px -55% 0px",
        threshold: [0, 0.15, 0.4, 0.75, 1],
      }
    );

    sections.forEach(function (section) {
      observer.observe(section);
    });
  }

  /* ------------------------------------------------------------ init */
  function init() {
    enhanceCodeBlocks();
    enhanceActiveLinks();

    // Reflect the resolved theme on the toggle's label (do not persist here,
    // so the OS preference keeps working until the user chooses explicitly).
    applyTheme(currentTheme(), false);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
