(function () {
  const REPO = "bwya77/W365-CLI-Native";
  const API_LATEST = `https://api.github.com/repos/${REPO}/releases/latest`;
  const FALLBACK_TAG = "v0.2.2";

  const ASSET_LABELS = {
    "w365-win-x64.zip": "Windows x64 (portable)",
    "w365-win-arm64.zip": "Windows ARM64 (portable)",
    "w365-osx-x64.zip": "macOS Intel",
    "w365-osx-arm64.zip": "macOS Apple Silicon",
  };

  function detectPlatform() {
    const ua = navigator.userAgent || "";
    const platform = navigator.platform || "";
    const isMac = /Mac/i.test(platform) || /Macintosh/i.test(ua);
    const isWin = /Win/i.test(platform) || /Windows/i.test(ua);

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

    if (isWin) return { os: "windows", arch, label: arch === "arm64" ? "Windows ARM64" : "Windows x64" };
    if (isMac) return { os: "macos", arch, label: arch === "arm64" ? "macOS (Apple Silicon)" : "macOS (Intel)" };
    return { os: "other", arch: "x64", label: "your platform" };
  }

  function zipAssetKeyFor(os, arch) {
    if (os === "windows") return arch === "arm64" ? "w365-win-arm64.zip" : "w365-win-x64.zip";
    if (os === "macos") return arch === "arm64" ? "w365-osx-arm64.zip" : "w365-osx-x64.zip";
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

    activate(defaultOs === "macos" ? "macos" : "windows");
  }

  async function init() {
    setYear();
    setupCopyButtons();

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

    // Windows defaults to the installer (easiest, one click, no terminal needed); macOS has no
    // double-click installer, so the primary action there is copying the install command instead
    // of downloading a file directly.
    if (detected.os === "windows") {
      const url = resolveInstaller(detected.arch);
      const label = `Download installer for ${detected.label}`;
      [[primaryBtn, primaryLabel], [primaryBtn2, primaryLabel2]].forEach(([btn, lbl]) => {
        if (!btn) return;
        btn.href = url;
        btn.onclick = null;
        if (lbl) lbl.textContent = label;
      });
      if (navBtn) { navBtn.href = url; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = `Detected ${detected.label}. Prefer the command line? Use the Windows tab below.`;
    } else if (detected.os === "macos") {
      const copyCommand = (evt) => {
        evt.preventDefault();
        const codeEl = document.getElementById("cmd-macos");
        const text = codeEl ? codeEl.textContent : "";
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).catch(() => {});
        }
        document.querySelector('[data-os-tab="macos"]')?.click();
        const btn = evt.currentTarget;
        const labelEl = btn.querySelector("span");
        const originalText = labelEl ? labelEl.textContent : btn.textContent;
        if (labelEl) labelEl.textContent = "Copied! Paste into Terminal";
        setTimeout(() => { if (labelEl) labelEl.textContent = originalText; }, 2200);
      };
      [primaryBtn, primaryBtn2].forEach((btn) => {
        if (!btn) return;
        btn.href = "#top";
        btn.onclick = copyCommand;
      });
      if (primaryLabel) primaryLabel.textContent = "Copy install command";
      if (primaryLabel2) primaryLabel2.textContent = "Copy install command";
      if (navBtn) { navBtn.href = "#top"; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Get started";
      platformNote.textContent = `Detected ${detected.label}. Paste the command below into Terminal.`;
    } else {
      [[primaryBtn, primaryLabel], [primaryBtn2, primaryLabel2]].forEach(([btn, lbl]) => {
        if (!btn) return;
        btn.href = `https://github.com/${REPO}/releases/latest`;
        btn.onclick = null;
        if (lbl) lbl.textContent = "See all downloads";
      });
      if (navBtn) { navBtn.href = `https://github.com/${REPO}/releases/latest`; navBtn.onclick = null; }
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = "Choose your platform below.";
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
