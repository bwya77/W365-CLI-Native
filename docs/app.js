(function () {
  const REPO = "bwya77/W365-CLI-Native";
  const API_LATEST = `https://api.github.com/repos/${REPO}/releases/latest`;
  const FALLBACK_TAG = "v0.1.5";

  const ASSET_LABELS = {
    "w365-win-x64.zip": "Windows x64",
    "w365-win-arm64.zip": "Windows ARM64",
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

  function assetKeyFor(os, arch) {
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

  async function init() {
    setYear();

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

    const primaryBtn = document.getElementById("primary-download");
    const primaryLabel = document.getElementById("primary-download-label");
    const primaryBtn2 = document.getElementById("primary-download-2");
    const primaryLabel2 = document.getElementById("primary-download-label-2");
    const navBtn = document.getElementById("nav-download");
    const navLabel = document.getElementById("nav-download-label");
    const platformNote = document.getElementById("platform-note");
    const versionPills = document.querySelectorAll("[data-version-pill]");
    const platformLinks = document.querySelectorAll("[data-asset]");

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

    const releaseBase = `https://github.com/${REPO}/releases/download/${tag}/`;
    const resolve = (name) => (assets && assets[name]) || releaseBase + name;

    // Update version pills
    versionPills.forEach((el) => { el.textContent = tag; });

    // Primary CTA
    const key = assetKeyFor(detected.os, detected.arch);
    if (key) {
      primaryBtn.href = resolve(key);
      primaryLabel.textContent = `Download for ${ASSET_LABELS[key]}`;
      if (primaryBtn2) primaryBtn2.href = resolve(key);
      if (primaryLabel2) primaryLabel2.textContent = `Download for ${ASSET_LABELS[key]}`;
      if (navBtn) navBtn.href = resolve(key);
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = `Detected ${detected.label}. Not right? Pick a build below.`;
    } else {
      primaryBtn.href = `https://github.com/${REPO}/releases/latest`;
      primaryLabel.textContent = "See all downloads";
      if (primaryBtn2) primaryBtn2.href = `https://github.com/${REPO}/releases/latest`;
      if (primaryLabel2) primaryLabel2.textContent = "See all downloads";
      if (navBtn) navBtn.href = `https://github.com/${REPO}/releases/latest`;
      if (navLabel) navLabel.textContent = "Download";
      platformNote.textContent = "Choose your platform below.";
    }

    // Platform link list
    platformLinks.forEach((a) => {
      const assetName = a.getAttribute("data-asset");
      a.href = resolve(assetName);
      if (assetName === key) a.classList.add("current");
    });

    const latestReleaseLink = document.getElementById("latest-release-link");
    if (latestReleaseLink) latestReleaseLink.href = `https://github.com/${REPO}/releases/latest`;
  }

  document.addEventListener("DOMContentLoaded", init);
})();
