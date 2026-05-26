/// <reference types="bun" />
import { marked } from "marked";

const sourcePath = new URL("../README.MD", import.meta.url);
const outputPath = new URL("../docs/index.html", import.meta.url);
const outputPathLabel = "docs/index.html";

function extractTitle(markdown: string): string {
  const headingMatch = markdown.match(/^#\s+(.+)$/m);
  if (headingMatch?.[1]) {
    return headingMatch[1].replace(/`/g, "").trim();
  }
  return "Project Documentation";
}

function rewriteRelativeUrl(url: string | null | undefined): string | null | undefined {
  if (!url || url.startsWith("#") || /^[a-z]+:/i.test(url) || url.startsWith("//")) {
    return url;
  }

  if (url.startsWith("./")) {
    return `../${url.slice(2)}`;
  }

  if (url.startsWith("/")) {
    return `..${url}`;
  }

  return `../${url}`;
}

function escapeHtml(text: string | null | undefined): string {
  return (text ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isVideoAsset(url: string | null | undefined): boolean {
  if (!url) {
    return false;
  }

  const cleanUrl = url.split(/[?#]/, 1)[0].toLowerCase();
  return cleanUrl.endsWith(".mp4") || cleanUrl.endsWith(".webm") || cleanUrl.endsWith(".ogg");
}

async function buildDocs(): Promise<void> {
  const markdown = await Bun.file(sourcePath).text();
  const title = extractTitle(markdown);
  let hasMermaidBlocks = false;

  marked.setOptions({
    gfm: true,
    breaks: false
  });

  const renderer = new marked.Renderer();
  const baseImageRenderer = renderer.image.bind(renderer);
  const baseLinkRenderer = renderer.link.bind(renderer);
  const baseCodeRenderer = renderer.code.bind(renderer);

  renderer.image = ({ href, title: imageTitle, text }) => {
    const rewrittenHref = rewriteRelativeUrl(href);

    if (!isVideoAsset(rewrittenHref)) {
      return baseImageRenderer({
        href: rewrittenHref,
        title: imageTitle,
        text
      });
    }

    const safeHref = escapeHtml(rewrittenHref);
    const safeTitle = escapeHtml(imageTitle || text || "Embedded video");
    const safeText = escapeHtml(text || "Video");
    return `<figure><video controls preload="metadata" title="${safeTitle}"><source src="${safeHref}">Your browser does not support embedded videos. <a href="${safeHref}">Open video</a>.</video><figcaption>${safeText}</figcaption></figure>`;
  };

  renderer.link = ({ href, title: linkTitle, text, tokens }) =>
    baseLinkRenderer({
      href: rewriteRelativeUrl(href),
      title: linkTitle,
      text,
      tokens
    });

  renderer.code = ({ text, lang, escaped }) => {
    if (lang?.toLowerCase() === "mermaid") {
      hasMermaidBlocks = true;
      return `<pre class="mermaid">${escapeHtml(text)}</pre>`;
    }

    return baseCodeRenderer({ text, lang, escaped });
  };

  const htmlBody = await marked.parse(markdown, { renderer });
  const mermaidScript = hasMermaidBlocks
    ? `
  <script type="module">
    import mermaid from "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs";
    const isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    mermaid.initialize({ startOnLoad: true, theme: isDark ? "dark" : "default" });
  </script>`
    : "";
  const fullHtml = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${title}</title>
  <style>
    :root {
      color-scheme: light dark;
    }

    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      line-height: 1.6;
      padding: 2rem 1rem;
    }

    main {
      max-width: 980px;
      margin: 0 auto;
    }

    img,
    video {
      max-width: 100%;
      height: auto;
    }

    pre {
      overflow-x: auto;
      padding: 1rem;
      border-radius: 8px;
      background: rgba(127, 127, 127, 0.12);
    }

    pre.mermaid {
      padding: 0;
      background: transparent;
      border-radius: 0;
      overflow: visible;
    }

    code {
      font-family: Consolas, "Courier New", monospace;
    }

    table {
      border-collapse: collapse;
      width: 100%;
    }

    th,
    td {
      border: 1px solid rgba(127, 127, 127, 0.3);
      padding: 0.4rem 0.6rem;
      text-align: left;
    }
  </style>
</head>
<body>
  <main>
${htmlBody}
  </main>
${mermaidScript}
</body>
</html>
`;

  await Bun.write(outputPath, fullHtml);
  console.log(`Built ${outputPathLabel}`);
}

buildDocs().catch((error) => {
  console.error(error);
  process.exit(1);
});
