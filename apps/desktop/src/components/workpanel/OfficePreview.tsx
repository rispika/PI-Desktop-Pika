import { lazy, Suspense, type CSSProperties } from "react";
import { useTranslation } from "react-i18next";
import type {
  OfficeParagraph,
  OfficePreview as OfficePreviewData,
  OfficeTextRun,
} from "@pi-desktop/shared";
import { WordDocumentPreview } from "./WordDocumentPreview";
import type { SpreadsheetEditor } from "./SpreadsheetPreview";

const SpreadsheetPreview = lazy(() => import("./SpreadsheetPreview").then(module => ({ default: module.SpreadsheetPreview })));

function runContent(run: OfficeTextRun, index: number) {
  let content: React.ReactNode = run.text;
  if (run.underline) content = <u>{content}</u>;
  if (run.italic) content = <em>{content}</em>;
  if (run.bold) content = <strong>{content}</strong>;
  return <span key={index} style={{ color: run.color, fontSize: run.fontSize ? `calc(${run.fontSize} * var(--office-slide-scale, 0.133333cqw))` : undefined }}>{content}</span>;
}

function Paragraph({ paragraph }: { paragraph: OfficeParagraph }) {
  const style = paragraph.style;
  if (!style) {
    return <p className="office-preview-paragraph" style={{ textAlign: paragraph.align }}>{paragraph.runs.map(runContent)}</p>;
  }
  const Heading = style === "title" ? "h1" : style === "heading1" ? "h2" : style === "heading2" ? "h3" : "h4";
  return <Heading className={`office-preview-${style}`}>{paragraph.runs.map(runContent)}</Heading>;
}

function PresentationPreview({
  data,
}: {
  data: Extract<OfficePreviewData, { format: "presentation" }>;
}) {
  const { t } = useTranslation();
  return (
    <div className="office-slides-preview">
      {data.slides.map((slide) => (
        <section className="office-slide" key={slide.number}>
          <div
            className="office-slide-canvas"
            style={{
              aspectRatio: `${data.slideWidth} / ${data.slideHeight}`,
              background: slide.background,
              "--office-slide-scale": `${1270000 / data.slideWidth}cqw`,
            } as CSSProperties}
          >
            <span className="office-slide-number">{slide.number}</span>
            {slide.elements?.length && slide.elements.flatMap(element => element.paragraphs).map(paragraph => paragraph.text).join("") === slide.paragraphs.map(paragraph => paragraph.text).join("") && slide.elements.filter(element => element.imageDataUrl).length === slide.imageDataUrls.length ? (
              slide.elements.map((element, index) => <div key={index} className="office-slide-element" style={{ left: `${element.x / data.slideWidth * 100}%`, top: `${element.y / data.slideHeight * 100}%`, width: `${element.width / data.slideWidth * 100}%`, height: `${element.height / data.slideHeight * 100}%`, transform: `rotate(${element.rotation ?? 0}deg)`, background: element.background }}>
                {element.imageDataUrl && <img src={element.imageDataUrl} alt="" />}
                {element.paragraphs.map((paragraph, paragraphIndex) => <Paragraph key={paragraphIndex} paragraph={paragraph} />)}
              </div>)
            ) : <div className="office-slide-content">
              {slide.paragraphs.map((paragraph, index) => (
                <Paragraph key={index} paragraph={paragraph} />
              ))}
              {slide.imageDataUrls.length > 0 && (
                <div className="office-slide-images">
                  {slide.imageDataUrls.map((src, index) => (
                    <img key={index} src={src} alt="" />
                  ))}
                </div>
              )}
            </div>}
          </div>
        </section>
      ))}
      {data.truncated && (
        <div className="office-preview-cap" aria-label={t("panel.files.previewTruncated")}>
          …
        </div>
      )}
    </div>
  );
}

export function OfficePreview({ preview, spreadsheetEditor }: { preview: OfficePreviewData; spreadsheetEditor?: SpreadsheetEditor }) {
  const { t } = useTranslation();
  if (preview.format === "word") {
    return (
      <WordDocumentPreview packageBase64={preview.packageBase64}><article className="office-word-preview">
        {preview.blocks.map((block, index) =>
          block.kind === "paragraph" ? (
            <Paragraph key={index} paragraph={block.paragraph} />
          ) : (
            <div className="office-word-table-wrap" key={index}>
              <table className="office-word-table">
                <tbody>
                  {block.rows.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {row.map((cell, columnIndex) => (
                        <td key={columnIndex}>{cell}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ),
        )}
        {preview.truncated && (
          <div className="office-preview-cap" aria-label={t("panel.files.previewTruncated")}>
            …
          </div>
        )}
      </article></WordDocumentPreview>
    );
  }
  if (preview.format === "spreadsheet") return (
    <Suspense fallback={<div className="office-preview-cap" role="status">{t("common.loading")}</div>}>
      <SpreadsheetPreview key={spreadsheetEditor?.identity ?? preview.revision ?? "readonly"} data={preview} editor={spreadsheetEditor} />
    </Suspense>
  );
  return <PresentationPreview data={preview} />;
}
