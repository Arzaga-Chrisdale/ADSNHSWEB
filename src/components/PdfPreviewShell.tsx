import {
  useState,
  useMemo,
  isValidElement,
  cloneElement,
  type ReactNode,
  type ReactElement,
  type CSSProperties,
} from "react";
import { ZoomIn, ZoomOut, RotateCw, Printer, ClipboardCopy, FileSpreadsheet } from "lucide-react";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { toast } from "sonner";

export const DEPED_BLUE = "#0038A8";
export const DEPED_RED = "#CE1126";
export const DEPED_YELLOW = "#FCD116";

export type PdfPaper = "short" | "long" | "a4";

type Props = {
  fileName: string;
  /** Element id of the printable region inside `children`. */
  printTargetId: string;
  /** Base file name (no extension) used by the fallback HTML Word exporter. */
  docBaseName?: string;
  /** Short/Long content toggle — omit to hide. */
  length?: "short" | "full";
  onLengthChange?: (value: "short" | "full") => void;
  /** Paper orientation for printing and preview aspect ratio. */
  orientation?: "portrait" | "landscape";
  pageLabel?: string;
  /** Optional controlled paper value. */
  paper?: PdfPaper;
  /** Receives paper changes from the toolbar. */
  onPaperChange?: (paper: PdfPaper) => void;
  /** Optional real .docx exporter supplied by a school-form page. */
  onCopyToWord?: () => void | Promise<void>;
  /** Hide the Copy to Word button when a form should not offer Word export. */
  hideCopyToWord?: boolean;
  /** Optional Excel exporter supplied by a school-form page. */
  onExportToExcel?: () => void | Promise<void>;
  /** CSS page margin used by the browser print dialog. */
  printMargin?: string;
  /** Makes the printable element fill exactly one selected paper page. */
  fitPrintToPage?: boolean;
  children: ReactNode;
};

const PAPER_DIMENSIONS: Record<PdfPaper, readonly [string, string]> = {
  short: ["8.5in", "11in"],
  long: ["8.5in", "13in"],
  a4: ["210mm", "297mm"],
};

function getPrintPageDimensions(paper: PdfPaper, orientation: "portrait" | "landscape") {
  const [portraitWidth, portraitHeight] = PAPER_DIMENSIONS[paper];

  return orientation === "landscape"
    ? { width: portraitHeight, height: portraitWidth }
    : { width: portraitWidth, height: portraitHeight };
}

export function PdfPreviewShell({
  fileName,
  printTargetId,
  docBaseName,
  length,
  onLengthChange,
  orientation = "landscape",
  pageLabel,
  paper: controlledPaper,
  onPaperChange,
  onCopyToWord,
  hideCopyToWord = false,
  onExportToExcel,
  printMargin = "0.4in",
  fitPrintToPage = false,
  children,
}: Props) {
  const [zoom, setZoom] = useState(90);
  const [internalPaper, setInternalPaper] = useState<PdfPaper>("long");
  const paper = controlledPaper ?? internalPaper;

  const setPaper = (nextPaper: PdfPaper) => {
    if (controlledPaper === undefined) {
      setInternalPaper(nextPaper);
    }
    onPaperChange?.(nextPaper);
  };

  const paperDimensions = useMemo(() => {
    const isLandscape = orientation === "landscape";

    const physicalInches: Record<PdfPaper, [number, number]> = {
      long: [8.5, 13],
      a4: [8.27, 11.69],
      short: [8.5, 11],
    };

    const [portraitWidth, portraitHeight] = physicalInches[paper];
    const [widthInches, heightInches] = isLandscape
      ? [portraitHeight, portraitWidth]
      : [portraitWidth, portraitHeight];

    const pixelsPerInch = 1250 / 13;

    return {
      width: Math.round(widthInches * pixelsPerInch),
      height: Math.round(heightInches * pixelsPerInch),
    };
  }, [paper, orientation]);

  const zoomScale = zoom / 100;

  // Reserve the real scaled size of the bondpaper so the preview container
  // can show both horizontal and vertical scrollbars when zoomed.
  const scaledPaperDimensions = useMemo(
    () => ({
      width: Math.ceil(paperDimensions.width * zoomScale),
      height: Math.ceil(paperDimensions.height * zoomScale),
    }),
    [paperDimensions.height, paperDimensions.width, zoomScale],
  );

  const print = async () => {
    const element = document.getElementById(printTargetId);
    if (!element) {
      toast.error("Nothing to print yet");
      return;
    }

    // Print inside the current document so the cloned SF5 keeps the app's
    // already-loaded Tailwind/Vite styles. A blank popup can print before its
    // copied stylesheets finish loading, which produces an unformatted PDF.
    const printRootId = "pdf-preview-print-root";
    const printStyleId = "pdf-preview-print-style";

    document.getElementById(printRootId)?.remove();
    document.getElementById(printStyleId)?.remove();

    const pageDimensions = getPrintPageDimensions(paper, orientation);
    const printRoot = document.createElement("div");
    const printableCopy = element.cloneNode(true) as HTMLElement;
    const printStyle = document.createElement("style");
    const originalTitle = document.title;

    printRoot.id = printRootId;
    printableCopy.setAttribute("data-pdf-print-target", "true");
    printRoot.appendChild(printableCopy);

    const fitToPageStyles = fitPrintToPage
      ? `
            [data-pdf-print-target="true"] {
              box-sizing: border-box !important;
              width: 100% !important;
              height: 100% !important;
              max-width: none !important;
              max-height: none !important;
              overflow: hidden !important;
              break-inside: avoid-page;
              page-break-inside: avoid;
            }
          `
      : `
            [data-pdf-print-target="true"] {
              box-sizing: border-box !important;
              width: auto !important;
              height: auto !important;
              max-width: 100% !important;
            }
          `;

    printStyle.id = printStyleId;
    printStyle.textContent = `
      @page {
        size: ${pageDimensions.width} ${pageDimensions.height};
        margin: ${printMargin};
      }

      @media screen {
        #${printRootId} {
          position: fixed !important;
          top: 0 !important;
          left: -100000px !important;
          width: ${pageDimensions.width} !important;
          height: ${pageDimensions.height} !important;
          overflow: hidden !important;
          visibility: hidden !important;
          pointer-events: none !important;
        }
      }

      @media print {
        html,
        body {
          width: 100% !important;
          height: 100% !important;
          margin: 0 !important;
          padding: 0 !important;
          background: #fff !important;
          overflow: visible !important;
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }

        body > *:not(#${printRootId}) {
          display: none !important;
        }

        #${printRootId} {
          display: block !important;
          position: static !important;
          box-sizing: border-box !important;
          width: 100% !important;
          height: ${fitPrintToPage ? "100%" : "auto"} !important;
          margin: 0 !important;
          padding: 0 !important;
          overflow: visible !important;
          background: #fff !important;
        }

        #${printRootId} table {
          border-collapse: collapse;
        }

        ${fitToPageStyles}
      }
    `;

    const cleanup = () => {
      printRoot.remove();
      printStyle.remove();
      document.title = originalTitle;
    };

    document.head.appendChild(printStyle);
    document.body.appendChild(printRoot);
    document.title = fileName;

    try {
      const images = Array.from(printableCopy.querySelectorAll("img"));

      await Promise.all([
        document.fonts?.ready ?? Promise.resolve(),
        ...images.map((image) => image.decode?.().catch(() => undefined) ?? Promise.resolve()),
      ]);

      await new Promise<void>((resolve) =>
        window.requestAnimationFrame(() => window.requestAnimationFrame(() => resolve())),
      );

      window.print();
    } catch {
      toast.error("Unable to open the print dialog");
    } finally {
      cleanup();
    }
  };

  /**
   * Fallback for forms that have not yet implemented a real .docx exporter.
   * GSA supplies `onCopyToWord`, so this fallback is not used for GSA.
   */
  const copyWordFallback = () => {
    const element = document.getElementById(printTargetId);
    if (!element) {
      toast.error("Nothing to copy yet");
      return;
    }

    const html = `<html
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40"
    >
      <head>
        <meta charset="utf-8" />
        <title>${fileName}</title>
      </head>
      <body>${element.outerHTML}</body>
    </html>`;

    const blob = new Blob([html], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");

    anchor.href = url;
    anchor.download = `${docBaseName || fileName.replace(/\.[a-z]+$/i, "")}.doc`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();

    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast.success("Downloaded as Word document");
  };

  const handleCopyToWord = () => {
    if (onCopyToWord) {
      void onCopyToWord();
      return;
    }

    copyWordFallback();
  };

  const handleExportToExcel = () => {
    if (onExportToExcel) {
      void onExportToExcel();
    }
  };

  let finalChildren = children;

  if (isValidElement(children)) {
    const element = children as ReactElement<{ style?: CSSProperties }>;

    finalChildren = cloneElement(element, {
      style: {
        ...(element.props.style || {}),
        width: paperDimensions.width,
        height: paperDimensions.height,
      },
    });
  }

  return (
    <div className="overflow-hidden rounded-2xl border shadow-sm print:border-0 print:shadow-none">
      <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[var(--admin-border)] bg-[var(--admin-white)] px-4 py-2 text-[var(--admin-text-heading)] print:hidden">
        <div className="flex items-center gap-2">
          <span className="text-sm font-medium">{fileName}</span>

          {length && onLengthChange && (
            <Select
              value={length}
              onValueChange={(value) => onLengthChange(value as "short" | "full")}
            >
              <SelectTrigger className="h-7 w-28 border-[var(--admin-input-border)] bg-[var(--admin-white)] text-xs text-[var(--admin-text-heading)]">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="short">Short</SelectItem>
                <SelectItem value="full">Full / Long</SelectItem>
              </SelectContent>
            </Select>
          )}

          <Select
            value={paper}
            onValueChange={(value) => setPaper(value as PdfPaper)}
            disabled={controlledPaper !== undefined && !onPaperChange}
          >
            <SelectTrigger className="h-7 w-28 border-[var(--admin-input-border)] bg-[var(--admin-white)] text-xs text-[var(--admin-text-heading)]">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="long">Long (8.5×13)</SelectItem>
              <SelectItem value="short">Short (Letter)</SelectItem>
              <SelectItem value="a4">A4</SelectItem>
            </SelectContent>
          </Select>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            type="button"
            onClick={() => setZoom((current) => Math.max(50, current - 10))}
            className="rounded-md p-1.5 hover:bg-[var(--admin-color-faf3ed)]"
            aria-label="Zoom out"
          >
            <ZoomOut className="size-4" />
          </button>

          <span className="w-10 text-center text-xs font-medium tabular-nums">{zoom}%</span>

          <button
            type="button"
            onClick={() => setZoom((current) => Math.min(200, current + 10))}
            className="rounded-md p-1.5 hover:bg-[var(--admin-color-faf3ed)]"
            aria-label="Zoom in"
          >
            <ZoomIn className="size-4" />
          </button>

          <button
            type="button"
            onClick={() => setZoom(90)}
            className="rounded-md p-1.5 hover:bg-[var(--admin-color-faf3ed)]"
            aria-label="Reset"
          >
            <RotateCw className="size-4" />
          </button>

          <button
            type="button"
            onClick={print}
            className="inline-flex items-center gap-1.5 rounded-md border bg-slate-800 px-3 py-1.5 text-xs font-medium text-white hover:bg-slate-700"
          >
            <Printer className="size-3.5" />
            Print / Save as PDF
          </button>

          {!hideCopyToWord && (
            <button
              type="button"
              onClick={handleCopyToWord}
              className="inline-flex items-center gap-1.5 rounded-md px-3 py-1.5 text-xs font-bold text-white shadow"
              style={{
                background: `linear-gradient(135deg, ${DEPED_BLUE}, #1d4ed8)`,
              }}
            >
              <ClipboardCopy className="size-3.5" />
              COPY TO WORD
            </button>
          )}

          {onExportToExcel && (
            <button
              type="button"
              onClick={handleExportToExcel}
              className="inline-flex items-center gap-1.5 rounded-md bg-emerald-600 px-3 py-1.5 text-xs font-bold text-white shadow hover:bg-emerald-700"
            >
              <FileSpreadsheet className="size-3.5" />
              EXPORT TO EXCEL
            </button>
          )}

          {pageLabel && (
            <span className="ml-2 text-xs text-[var(--admin-text-muted)]">{pageLabel}</span>
          )}
        </div>
      </div>

      <div
        className="max-h-[calc(100vh-230px)] min-h-[430px] overflow-auto bg-[var(--admin-color-ece9e6)] print:max-h-none print:min-h-0 print:overflow-visible print:bg-white"
        style={{
          scrollbarGutter: "stable both-edges",
          overscrollBehavior: "contain",
        }}
      >
        <div className="flex min-w-full w-max justify-center p-4 print:block print:p-0 md:p-8">
          <div
            className="relative shrink-0 print:h-auto print:w-auto"
            style={{
              width: scaledPaperDimensions.width,
              height: scaledPaperDimensions.height,
            }}
          >
            <div
              className="absolute left-0 top-0 origin-top-left bg-white shadow-2xl print:static print:transform-none print:shadow-none"
              style={{
                transform: `scale(${zoomScale})`,
                transformOrigin: "top left",
              }}
            >
              {finalChildren}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
