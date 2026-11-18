import { useEffect, useRef, useState, type ReactNode } from "react";
import { useTranslation } from "react-i18next";

const FRAME_DOCUMENT = `<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src blob: data:; style-src 'unsafe-inline'; font-src data: blob:; base-uri 'none'; form-action 'none'"><style>html,body{margin:0;background:#eee}body{overflow:auto}.docx-wrapper{padding:16px!important}section.docx{margin:0 auto 16px!important}</style></head><body></body></html>`;

export function WordDocumentPreview({ packageBase64, children }: { packageBase64?: string; children: ReactNode }) {
  const { t } = useTranslation();
  const container = useRef<HTMLDivElement>(null);
  const [renderedPackage, setRenderedPackage] = useState<string>();
  useEffect(() => {
    if (!packageBase64 || !container.current) return;
    let cancelled = false;
    setRenderedPackage(undefined);
    const frame = document.createElement("iframe");
    frame.className = "office-word-frame";
    frame.title = t("panel.files.preview");
    frame.setAttribute("sandbox", "allow-same-origin");
    frame.referrerPolicy = "no-referrer";
    frame.onload = async () => {
      try {
        const [{ renderAsync }, { default: DOMPurify }] = await Promise.all([import("docx-preview"), import("dompurify")]);
        if (cancelled) return;
        const doc = frame.contentDocument;
        if (!doc) throw new Error("Word preview frame is unavailable");
        const bytes = Uint8Array.from(atob(packageBase64), value => value.charCodeAt(0));
        await renderAsync(bytes, doc.body, doc.head, {
          renderAltChunks: false,
          useBase64URL: true,
          renderComments: false,
          renderChanges: false,
          breakPages: true,
        });
        if (cancelled) return;
        if (!doc.querySelector('meta[http-equiv="Content-Security-Policy"]')) {
          const policy = doc.createElement("meta");
          policy.httpEquiv = "Content-Security-Policy";
          policy.content = "default-src 'none'; img-src blob: data:; style-src 'unsafe-inline'; font-src data: blob:; base-uri 'none'; form-action 'none'";
          doc.head.prepend(policy);
        }
        DOMPurify.sanitize(doc.body, { IN_PLACE: true, ADD_TAGS: ["style"], FORBID_TAGS: ["iframe", "object", "embed", "form", "input", "button", "svg", "math"] });
        // Keep links inert even when the source document contains internal hyperlinks.
        for (const anchor of doc.querySelectorAll("a")) anchor.removeAttribute("href");
        setRenderedPackage(packageBase64);
      } catch (error) {
        if (!cancelled) {
          console.warn("Word layout preview failed; using text preview", error instanceof Error ? `${error.name}: ${error.message}\n${error.stack ?? ""}` : error);
          frame.remove();
        }
      }
    };
    frame.srcdoc = FRAME_DOCUMENT;
    container.current.append(frame);
    return () => {
      cancelled = true;
      frame.onload = null;
      frame.remove();
    };
  }, [packageBase64, t]);
  const rendered = Boolean(packageBase64 && renderedPackage === packageBase64);
  return <><div ref={container} className="office-word-layout" hidden={!rendered} />{!rendered && children}</>;
}
