const formatSize = (bytes) => `${(bytes / 1024 / 1024).toFixed(1)} MB`;

async function loadRelease() {
  try {
    const response = await fetch("/api/v1/releases/current", { cache: "no-store" });
    if (!response.ok) return;
    const release = await response.json();
    document.querySelectorAll("[data-release-version]").forEach((node) => {
      node.textContent = `v${release.version}`;
    });
    document.querySelectorAll("[data-release-size]").forEach((node) => {
      node.textContent = `Windows 64 位 · ${formatSize(release.installer.size)}`;
    });
    document.querySelectorAll("[data-footer-version]").forEach((node) => {
      node.textContent = `当前版本 v${release.version}`;
    });
    document.querySelector("[data-changelog-version]").textContent = `v${release.version} 功能更新`;
    document.querySelector("[data-release-date]").textContent = new Date(release.publishedAt).toLocaleDateString("zh-CN", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    const notes = document.querySelector("[data-release-notes]");
    notes.replaceChildren(...release.releaseNotes.map((note) => {
      const item = document.createElement("li");
      item.textContent = note;
      return item;
    }));
  } catch {
    // Keep the static fallback copy when the release API is temporarily unavailable.
  }
}

void loadRelease();
