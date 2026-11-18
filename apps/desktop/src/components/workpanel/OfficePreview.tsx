import { useEffect, useState, type CSSProperties } from "react";
import { useTranslation } from "react-i18next";
import type {
  OfficeParagraph,
  OfficePreview as OfficePreviewData,
  OfficeTextRun,
} from "@pi-desktop/shared";
import { Button, cx } from "../ui";
import { WordDocumentPreview } from "./WordDocumentPreview";

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

function SpreadsheetPreview({
  data,
}: {
  data: Extract<OfficePreviewData, { format: "spreadsheet" }>;
}) {
  const { t } = useTranslation();
  const [active, setActive] = useState(0);
  useEffect(() => {
    setActive((current) => Math.min(current, Math.max(0, data.sheets.length - 1)));
  }, [data.sheets.length]);
  const sheet = data.sheets[active];
  const columns = Math.max(0, ...(sheet?.rows.map(row => row.length) ?? []));
  const columnName = (index: number): string => index < 26 ? String.fromCharCode(65 + index) : columnName(Math.floor(index / 26) - 1) + columnName(index % 26);
  if (!sheet) return <div className="office-preview-empty">—</div>;

  return (
    <div className="office-sheet-preview">
      <div className="office-sheet-tabs" role="tablist">
        {data.sheets.map((item, index) => (
          <Button
            key={`${item.name}:${index}`}
            type="button"
            role="tab"
            size="sm"
            variant="ghost"
            aria-selected={index === active}
            className={cx("office-sheet-tab", index === active && "selected")}
            onClick={() => setActive(index)}
          >
            {item.name}
          </Button>
        ))}
      </div>
      <div className="office-sheet-scroll">
        <table className="office-sheet-table">
          <thead><tr><th />{Array.from({ length: columns }, (_, index) => <th scope="col" key={index}>{columnName(index)}</th>)}</tr></thead>
          <tbody>
            {sheet.rows.map((row, rowIndex) => (
              <tr key={rowIndex}>
                <th scope="row">{rowIndex + 1}</th>
                {Array.from({ length: columns }, (_, columnIndex) => (
                  <td key={columnIndex}>{row[columnIndex] ?? ""}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {(sheet.truncated || data.truncated) && (
        <div className="office-preview-cap" aria-label={t("panel.files.previewTruncated")}>
          …
        </div>
      )}
    </div>
  );
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

export function OfficePreview({ preview }: { preview: OfficePreviewData }) {
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
  if (preview.format === "spreadsheet") return <SpreadsheetPreview data={preview} />;
  return <PresentationPreview data={preview} />;
}
