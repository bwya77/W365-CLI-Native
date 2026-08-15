(function () {
  const REPO = "bwya77/W365-CLI-Native";
  const API_LATEST = `https://api.github.com/repos/${REPO}/releases/latest`;
  const FALLBACK_TAG = "v0.5.32";

  const ASSET_LABELS = {
    "w365-win-x64.zip": "Windows x64 (portable)",
    "w365-win-arm64.zip": "Windows ARM64 (portable)",
    "w365-osx-x64.zip": "macOS Intel",
    "w365-osx-arm64.zip": "macOS Apple Silicon",
    "w365-linux-x64.tar.gz": "Linux x64",
    "w365-linux-arm64.tar.gz": "Linux ARM64",
  };

  function detectPlatform() {
    const ua = navigator.userAgent || "";
    const platform = navigator.platform || "";
    const isMac = /Mac/i.test(platform) || /Macintosh/i.test(ua);
    const isWin = /Win/i.test(platform) || /Windows/i.test(ua);
    // Android UAs also contain "Linux", so explicitly exclude Android/mobile before matching —
    // otherwise every Android visitor would be mis-detected as a desktop Linux user.
    const isLinux = !isMac && !isWin && (/Linux/i.test(platform) || /Linux/i.test(ua)) && !/Android/i.test(ua);

    // Best-effort arch detection. Apple Silicon Macs still often report "MacIntel"
    // so we check for indicators of Rosetta/ARM where possible.
    let arch = "x64";
    if (isMac) {
      // No reliable client-side ARM64 signal on Safari; assume Apple Silicon by default
      // since it has shipped standard since 2020, but respect a maximum touch heuristic.
      arch = "arm64";
    }
    if (isWin && /ARM/i.test(ua)) {
      arch = "arm64";
    }
    if (isLinux && /aarch64|arm64/i.test(ua)) {
      arch = "arm64";
    }

    if (isWin) return { os: "windows", arch, label: arch === "arm64" ? "Windows ARM64" : "Windows x64" };
    if (isMac) return { os: "macos", arch, label: arch === "arm64" ? "macOS (Apple Silicon)" : "macOS (Intel)" };
    if (isLinux) return { os: "linux", arch, label: arch === "arm64" ? "Linux ARM64" : "Linux x64" };
    return { os: "other", arch: "x64", label: "your platform" };
  }

  function zipAssetKeyFor(os, arch) {
    if (os === "windows") return arch === "arm64" ? "w365-win-arm64.zip" : "w365-win-x64.zip";
    if (os === "macos") return arch === "arm64" ? "w365-osx-arm64.zip" : "w365-osx-x64.zip";
    if (os === "linux") return arch === "arm64" ? "w365-linux-arm64.tar.gz" : "w365-linux-x64.tar.gz";
    return null;
  }

  async function getHighEntropyArch() {
    try {
      if (navigator.userAgentData && navigator.userAgentData.getHighEntropyValues) {
        const info = await navigator.userAgentData.getHighEntropyValues(["architecture", "bitness"]);
        if (info.architecture) {
          if (/arm/i.test(info.architecture)) return "arm64";
          return "x64";
        }
      }
    } catch (e) { /* ignore, fall back */ }
    return null;
  }

  function setYear() {
    const el = document.getElementById("year");
    if (el) el.textContent = new Date().getFullYear();
  }

  // Windows installer filenames embed the version (e.g. W365CLISetup-0.2.2-win-x64.exe), so they
  // can't be looked up by a fixed asset name the way the zips can. Find them by pattern instead,
  // either from the fetched release asset list or (offline fallback) by constructing the expected
  // filename from the fallback tag's version number.
  function findInstallerAsset(assets, arch) {
    const suffix = `-win-${arch}.exe`;
    if (assets) {
      const match = Object.keys(assets).find((name) => name.startsWith("W365CLISetup-") && name.endsWith(suffix));
      if (match) return { name: match, url: assets[match] };
    }
    return null;
  }

  function setupCopyButtons() {
    document.querySelectorAll(".copy-btn").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const targetId = btn.getAttribute("data-copy-target");
        const codeEl = document.getElementById(targetId);
        if (!codeEl) return;
        const text = codeEl.textContent;

        try {
          if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(text);
          } else {
            const textarea = document.createElement("textarea");
            textarea.value = text;
            textarea.style.position = "fixed";
            textarea.style.opacity = "0";
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand("copy");
            document.body.removeChild(textarea);
          }
          const original = btn.textContent;
          btn.textContent = "Copied!";
          btn.classList.add("copied");
          setTimeout(() => {
            btn.textContent = original;
            btn.classList.remove("copied");
          }, 1800);
        } catch (e) { /* clipboard denied; user can still select the text manually */ }
      });
    });
  }

  function setupCmdTabs(defaultOs) {
    const tabs = document.querySelectorAll(".cmd-tab");
    const panels = document.querySelectorAll(".cmd-panel");

    function activate(os) {
      tabs.forEach((t) => t.classList.toggle("active", t.getAttribute("data-os-tab") === os));
      panels.forEach((p) => p.classList.toggle("active", p.getAttribute("data-os-panel") === os));
    }

    tabs.forEach((t) => {
      t.addEventListener("click", () => activate(t.getAttribute("data-os-tab")));
    });

    activate(defaultOs === "macos" || defaultOs === "linux" ? defaultOs : "windows");
  }

  function setupCmdToggle() {
    const toggle = document.getElementById("cmd-toggle");
    const box = document.getElementById("cmd-box");
    if (!toggle || !box) return null;

    function setOpen(open) {
      box.hidden = !open;
      toggle.classList.toggle("open", open);
      toggle.firstChild.textContent = open ? "Hide install command " : "Prefer the command line? ";
    }

    toggle.addEventListener("click", () => setOpen(box.hidden));
    return setOpen;
  }

  function setupDownloadMenu() {
    const toggle = document.getElementById("download-menu-toggle");
    const menu = document.getElementById("download-menu");
    if (!toggle || !menu) return;

    function setOpen(open) {
      menu.hidden = !open;
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
    }

    // stopPropagation here is what keeps a caret click from also triggering the parent button's
    // own click handler (which downloads/copies) — the caret is nested inside that button so its
    // clicks bubble up to it by default.
    toggle.addEventListener("click", (evt) => {
      evt.stopPropagation();
      setOpen(menu.hidden);
    });
    toggle.addEventListener("keydown", (evt) => {
      if (evt.key === "Enter" || evt.key === " ") {
        evt.preventDefault();
        evt.stopPropagation();
        setOpen(menu.hidden);
      }
    });

    // Close the menu on outside click, Escape, or after picking an item.
    document.addEventListener("click", (evt) => {
      if (!menu.hidden && !menu.contains(evt.target) && evt.target !== toggle) setOpen(false);
    });
    document.addEventListener("keydown", (evt) => {
      if (evt.key === "Escape" && !menu.hidden) setOpen(false);
    });
    menu.querySelectorAll("a").forEach((a) => a.addEventListener("click", () => setOpen(false)));
  }

  function setupHeroSlideshow() {
    const root = document.getElementById("hero-slideshow");
    if (!root) return;

    const slides = Array.from(root.querySelectorAll(".slide"));
    const dotsWrap = document.getElementById("slide-dots");
    const titleEl = document.getElementById("slide-title");
    const prevBtn = document.getElementById("slide-prev");
    const nextBtn = document.getElementById("slide-next");
    if (!slides.length) return;

    let index = Math.max(0, slides.findIndex((s) => s.classList.contains("is-active")));
    if (index < 0) index = 0;
    let timer = null;
    const AUTO_MS = 4500;

    const dots = slides.map((_, i) => {
      const d = document.createElement("button");
      d.type = "button";
      d.className = "slide-dot";
      d.setAttribute("aria-label", "Show screenshot " + (i + 1));
      d.addEventListener("click", () => show(i, true));
      dotsWrap.appendChild(d);
      return d;
    });

    function show(next, userTriggered) {
      index = ((next % slides.length) + slides.length) % slides.length;
      slides.forEach((s, i) => s.classList.toggle("is-active", i === index));
      dots.forEach((d, i) => d.classList.toggle("is-active", i === index));
      if (titleEl) titleEl.textContent = slides[index].getAttribute("data-title") || "W365 CLI";
      if (userTriggered) restart();
    }

    function restart() {
      if (timer) clearInterval(timer);
      timer = setInterval(() => show(index + 1, false), AUTO_MS);
    }

    if (prevBtn) prevBtn.addEventListener("click", () => show(index - 1, true));
    if (nextBtn) nextBtn.addEventListener("click", () => show(index + 1, true));
    root.addEventListener("mouseenter", () => { if (timer) clearInterval(timer); });
    root.addEventListener("mouseleave", restart);

    show(index, false);
    restart();
  }

  async function init() {
    setYear();
    setupCopyButtons();
    setupHeroSlideshow();

    const detected = detectPlatform();
    const highEntropyArch = await getHighEntropyArch();
    if (highEntropyArch) {
      detected.arch = highEntropyArch;
      if (detected.os === "windows") {
        detected.label = highEntropyArch === "arm64" ? "Windows ARM64" : "Windows x64";
      } else if (detected.os === "macos") {
        detected.label = highEntropyArch === "arm64" ? "macOS (Apple Silicon)" : "macOS (Intel)";
      }
    }

    setupCmdTabs(detected.os);
    const setCmdBoxOpen = setupCmdToggle();
    setupDownloadMenu();

    const primaryBtn = document.getElementById("primary-download");
    const primaryLabel = document.getElementById("primary-download-label");
    const primaryBtn2 = document.getElementById("primary-download-2");
    const primaryLabel2 = document.getElementById("primary-download-label-2");
    const navBtn = document.getElementById("nav-download");
    const navLabel = document.getElementById("nav-download-label");
    const platformNote = document.getElementById("platform-note");
    const versionPills = document.querySelectorAll("[data-version-pill]");
    const platformLinks = document.querySelectorAll("[data-asset], [data-asset-pattern]");

    let tag = FALLBACK_TAG;
    let assets = null;

    try {
      const res = await fetch(API_LATEST, { headers: { Accept: "application/vnd.github+json" } });
      if (res.ok) {
        const data = await res.json();
        tag = data.tag_name || tag;
        assets = {};
        (data.assets || []).forEach((a) => { assets[a.name] = a.browser_download_url; });
      }
    } catch (e) { /* offline or rate-limited: use fallback below */ }

    const version = tag.replace(/^v/, "");
    const releaseBase = `https://github.com/${REPO}/releases/download/${tag}/`;
    const resolveZip = (name) => (assets && assets[name]) || releaseBase + name;
    const resolveInstaller = (arch) => {
      const found = findInstallerAsset(assets, arch);
      if (found) return found.url;
      return releaseBase + `W365CLISetup-${version}-win-${arch}.exe`;
    };

    // Update version pills
    versionPills.forEach((el) => { el.textContent = tag; });

    // primary-download (hero) is a real <button>, not a link, because it embeds the caret
    // hot-zone as a child element — the button's own click handler (attached once, below) reads
    // these data attributes at click time to decide what to do. primary-download-2 (final CTA)
    // has no caret, so it stays a plain anchor/onclick as before.
    function copyShellCommandAndReveal(osTab) {
      const codeEl = document.getElementById(`cmd-${osTab}`);
      const text = codeEl ? codeEl.textContent : "";
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).catch(() => {});
      }
      if (setCmdBoxOpen) setCmdBoxOpen(true);
      document.querySelector(`[data-os-tab="${osTab}"]`)?.click();
      if (primaryLabel) {
        const original = primaryLabel.textContent;
        primaryLabel.textContent = "Copied! Paste into Terminal";
        setTimeout(() => { primaryLabel.textContent = original; }, 2200);
      }
    }

    // Windows defaults to the installer (easiest, one click, no terminal needed); macOS and
    // Linux have no double-click installer, so the primary action there is copying the install
    // command instead of downloading a file directly.
    if (detected.os === "windows") {
      const url = resolveInstaller(detected.arch);
      const label = `Download (${detected.arch === "arm64" ? "ARM64" : "x64"})`;
      if (primaryBtn) {
        primaryBtn.dataset.action = "";
        primaryBtn.dataset.href = url;
      }
      if (primaryLabel) primaryLabel.textContent = label;
      if (primaryBtn2) {
        primaryBtn2.href = url;
        primaryBtn2.onclick = null;
        if (primaryLabel2) primaryLabel2.textContent = label;
      }
      if (navBtn) { navBtn.href = url; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = `Detected ${detected.label}. Need a different variant? Use the arrow on the button.`;
    } else if (detected.os === "macos" || detected.os === "linux") {
      const osTab = detected.os;
      if (primaryBtn) {
        primaryBtn.dataset.action = "copy-shell-command";
        primaryBtn.dataset.osTab = osTab;
        primaryBtn.dataset.href = "";
      }
      if (primaryLabel) primaryLabel.textContent = "Copy install command";
      const copyCommandForLink = (evt) => {
        evt.preventDefault();
        copyShellCommandAndReveal(osTab);
      };
      if (primaryBtn2) {
        primaryBtn2.href = "#top";
        primaryBtn2.onclick = copyCommandForLink;
        if (primaryLabel2) primaryLabel2.textContent = "Copy install command";
      }
      if (navBtn) { navBtn.href = "#top"; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Get started";
      const terminalName = detected.os === "macos" ? "Terminal" : "your terminal";
      platformNote.textContent = `Detected ${detected.label}. Paste the command below into ${terminalName}.`;
    } else {
      const releasePage = `https://github.com/${REPO}/releases/latest`;
      if (primaryBtn) {
        primaryBtn.dataset.action = "";
        primaryBtn.dataset.href = releasePage;
      }
      if (primaryLabel) primaryLabel.textContent = "See all downloads";
      if (primaryBtn2) {
        primaryBtn2.href = releasePage;
        primaryBtn2.onclick = null;
        if (primaryLabel2) primaryLabel2.textContent = "See all downloads";
      }
      if (navBtn) { navBtn.href = releasePage; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = "Choose your platform below.";
    }

    // Single click handler for the hero button itself. Clicks that originated on the caret never
    // reach here (it calls stopPropagation), so this only fires for genuine "download" clicks.
    if (primaryBtn) {
      primaryBtn.addEventListener("click", () => {
        if (primaryBtn.dataset.action === "copy-shell-command") {
          copyShellCommandAndReveal(primaryBtn.dataset.osTab || "macos");
        } else if (primaryBtn.dataset.href) {
          window.location.href = primaryBtn.dataset.href;
        }
      });
    }

    // Direct asset link list (installers + portable zips)
    platformLinks.forEach((a) => {
      const zipName = a.getAttribute("data-asset");
      const pattern = a.getAttribute("data-asset-pattern");
      if (zipName) {
        a.href = resolveZip(zipName);
        if (zipName === zipAssetKeyFor(detected.os, detected.arch)) a.classList.add("current");
      } else if (pattern === "installer-x64") {
        a.href = resolveInstaller("x64");
        if (detected.os === "windows" && detected.arch === "x64") a.classList.add("current");
      } else if (pattern === "installer-arm64") {
        a.href = resolveInstaller("arm64");
        if (detected.os === "windows" && detected.arch === "arm64") a.classList.add("current");
      }
    });

    const latestReleaseLink = document.getElementById("latest-release-link");
    if (latestReleaseLink) latestReleaseLink.href = `https://github.com/${REPO}/releases/latest`;
  }

  document.addEventListener("DOMContentLoaded", init);
})();
