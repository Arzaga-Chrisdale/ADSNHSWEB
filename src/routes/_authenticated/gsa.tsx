import { Fragment } from "react";
import { createFileRoute, Link } from "@tanstack/react-router";
import { useQuery } from "@tanstack/react-query";
import { supabase } from "@/integrations/supabase/client";
import { useMemo, useState } from "react";
import { ArrowLeft, FileText, GraduationCap, Users } from "lucide-react";
import { computeAverage, type ClassRow, type GradeRow, type StudentRow } from "@/lib/data";
import { PdfPreviewShell, type PdfPaper } from "@/components/PdfPreviewShell";
import { DepEdHeader, DEPED_BLUE, DEPED_YELLOW } from "@/components/DepEdHeader";
import deped from "@/assets/deped.jpg";
import {
  AlignmentType,
  BorderStyle,
  Document as WordDocument,
  HeightRule,
  ImageRun,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

const SCHOOL_FORMS_ACTIVE_CLASS_KEY = "school-forms-active-class-id";

function readSchoolFormsClassId() {
  if (typeof window === "undefined") return "";
  return window.sessionStorage.getItem(SCHOOL_FORMS_ACTIVE_CLASS_KEY) ?? "";
}

export const Route = createFileRoute("/_authenticated/gsa")({
  component: GSAPage,
});

type Term = "1" | "2" | "3" | "final";
const TERM_LABEL: Record<Term, string> = { "1": "1st", "2": "2nd", "3": "3rd", final: "Final" };

type SubjectRow = {
  subject: string;
  strand: string;
  section: string;
  mReg: number; fReg: number;
  mPass: number; fPass: number;
  mSum: number; fSum: number;
};

function emptyRow(): SubjectRow {
  return { subject: "", strand: "N/A", section: "", mReg: 0, fReg: 0, mPass: 0, fPass: 0, mSum: 0, fSum: 0 };
}

function GSAPage() {
  const [classId, setClassId] = useState<string>(readSchoolFormsClassId);
  const [term, setTerm] = useState<Term>("1");
  const [paper, setPaper] = useState<PdfPaper>("long");
  const [teacherOverride, setTeacherOverride] = useState<string>("");
  const [signatories, setSignatories] = useState({
    prepared: "Teacher Name", preparedTitle: "Teacher I",
    checked: "Checker Name", checkedTitle: "Master Teacher I",
    noted: "Noted By Name", notedTitle: "SHS Dept. Head",
    approved: "Principal Name", approvedTitle: "School Principal",
  });

  const { data: profile } = useQuery({
    queryKey: ["profile"],
    queryFn: async () => (await supabase.from("profiles").select("*").maybeSingle()).data as any,
  });
  const { data: classes = [] } = useQuery({
    queryKey: ["classes"],
    queryFn: async () => (await supabase.from("classes").select("*").order("created_at", { ascending: false })).data as ClassRow[],
  });
  const selectedSchoolFormsClass = classes.find((item) => item.id === classId);
  const active = selectedSchoolFormsClass?.id || classes[0]?.id;
  const klass = classes.find((c) => c.id === active);

  const { data: students = [] } = useQuery({
    enabled: !!active,
    queryKey: ["students", active],
    queryFn: async () => (await supabase.from("students").select("*").eq("class_id", active!).order("last_name")).data as StudentRow[],
  });
  const { data: grades = [] } = useQuery({
    enabled: !!active,
    queryKey: ["grades-all", active],
    queryFn: async () => (await supabase.from("grades").select("*").eq("class_id", active!)).data as GradeRow[],
  });

  // Derive rows: one per unique subject taught to this class.
  const rows = useMemo<SubjectRow[]>(() => {
    if (!klass) return [emptyRow()];
    const subjects = Array.from(new Set([klass.subject, ...grades.map((g) => g.subject)])).filter(Boolean);
    if (subjects.length === 0) subjects.push(klass.subject || "Subject");
    return subjects.map((subj) => {
      const r = emptyRow();
      r.subject = subj;
      r.section = klass.section || "";
      for (const s of students) {
        const isM = s.sex === "male";
        const isF = s.sex === "female";
        if (isM) r.mReg++;
        if (isF) r.fReg++;
        const score = term === "final"
          ? computeAverage(["1", "2", "3"].map((t) => grades.find((g) => g.student_id === s.id && g.subject === subj && g.term === t)?.score ?? null))
          : (grades.find((g) => g.student_id === s.id && g.subject === subj && g.term === term)?.score ?? null);
        if (score == null) continue;
        if (isM) { r.mSum += score; if (score >= 75) r.mPass++; }
        if (isF) { r.fSum += score; if (score >= 75) r.fPass++; }
      }
      return r;
    });
  }, [students, grades, klass, term]);

  const reset = () => { setTerm("1"); setTeacherOverride(""); setClassId(""); };

  const teacherName = teacherOverride || klass?.teacher_name || profile?.full_name || "";

  const copyGsaToWord = async () => {
    if (!klass) {
      window.alert("Select a class before creating the Word document.");
      return;
    }

    try {
      const inchesToTwip = (inches: number) => Math.round(inches * 1440);

      // Pass PORTRAIT dimensions to docx. PageOrientation.LANDSCAPE swaps
      // them internally. Passing landscape dimensions here would swap them
      // a second time and create a narrow portrait page.
      const portraitPageInches =
        paper === "long"
          ? { width: 8.5, height: 13 }
          : paper === "short"
            ? { width: 8.5, height: 11 }
            : { width: 8.27, height: 11.69 };

      const marginInches = 0.3;
      const pageWidth = inchesToTwip(portraitPageInches.width);
      const pageHeight = inchesToTwip(portraitPageInches.height);
      const margin = inchesToTwip(marginInches);

      // In landscape mode the long portrait edge becomes the page width.
      const contentWidth = pageHeight - margin * 2;

      const weights = [10, 6, 7, ...Array.from({ length: 15 }, () => 77 / 15)];
      const columnWidths = weights.map((weight) =>
        Math.floor((contentWidth * weight) / 100),
      );

      // Put any rounding remainder into the final column so the table width
      // exactly matches the printable page width.
      columnWidths[columnWidths.length - 1] +=
        contentWidth - columnWidths.reduce((total, width) => total + width, 0);

      const noBorder = {
        style: BorderStyle.NONE,
        size: 0,
        color: "FFFFFF",
      };

      const thinBorder = {
        style: BorderStyle.SINGLE,
        size: 4,
        color: "000000",
      };

      const allThinBorders = {
        top: thinBorder,
        bottom: thinBorder,
        left: thinBorder,
        right: thinBorder,
      };

      const noBorders = {
        top: noBorder,
        bottom: noBorder,
        left: noBorder,
        right: noBorder,
        insideHorizontal: noBorder,
        insideVertical: noBorder,
      };

      let logoBytes: Uint8Array | null = null;
      let logoType: "png" | "jpg" = "jpg";

      try {
        const response = await fetch(deped);
        if (response.ok) {
          const contentType = response.headers.get("content-type") || "";
          logoType = contentType.includes("png") ? "png" : "jpg";
          logoBytes = new Uint8Array(await response.arrayBuffer());
        }
      } catch (logoError) {
        console.warn("Unable to embed the DepEd logo in the Word file.", logoError);
      }

      const centerLines = (
        lines: string[],
        options: {
          bold?: boolean;
          size?: number;
          color?: string;
          alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
        } = {},
      ) =>
        new Paragraph({
          alignment: options.alignment ?? AlignmentType.CENTER,
          spacing: { before: 0, after: 0, line: 240 },
          children: lines.flatMap((line, index) => [
            new TextRun({
              text: line,
              break: index === 0 ? undefined : 1,
              bold: options.bold,
              size: options.size ?? 12,
              color: options.color,
              font: "Arial",
            }),
          ]),
        });

      const makeCell = ({
        text,
        width,
        bold = false,
        size = 12,
        alignment = AlignmentType.CENTER,
        rowSpan,
        columnSpan,
        fill,
      }: {
        text: string | string[];
        width: number;
        bold?: boolean;
        size?: number;
        alignment?: (typeof AlignmentType)[keyof typeof AlignmentType];
        rowSpan?: number;
        columnSpan?: number;
        fill?: string;
      }) =>
        new TableCell({
          width: { size: width, type: WidthType.DXA },
          rowSpan,
          columnSpan,
          verticalAlign: VerticalAlign.CENTER,
          borders: allThinBorders,
          shading: fill ? { fill } : undefined,
          margins: { top: 35, bottom: 35, left: 30, right: 30 },
          children: [
            centerLines(Array.isArray(text) ? text : [text], {
              bold,
              size,
              alignment,
            }),
          ],
        });

      // Three balanced columns keep the four school-heading lines exactly
      // centered on the page while the logo remains on the left.
      const schoolHeaderWidth = Math.floor(contentWidth * 0.6);

      // Increase this ratio to move only the DepEd logo farther right.
      // Because the left and right side columns stay equal, the school
      // heading remains centered on the bond paper.
      const schoolSideRatio = 0.38;
      const schoolSideWidth = Math.floor(
        schoolHeaderWidth * schoolSideRatio,
      );
      const schoolTextWidth = schoolHeaderWidth - schoolSideWidth * 2;

      const schoolHeader = new Table({
        alignment: AlignmentType.CENTER,
        width: {
          size: schoolHeaderWidth,
          type: WidthType.DXA,
        },
        columnWidths: [schoolSideWidth, schoolTextWidth, schoolSideWidth],
        layout: TableLayoutType.FIXED,
        borders: noBorders,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              new TableCell({
                width: {
                  size: schoolSideWidth,
                  type: WidthType.DXA,
                },
                borders: noBorders,
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    indent: {
                      right: 0,
                    },
                    children:
                      logoBytes !== null
                        ? [
                            new ImageRun({
                              data: logoBytes,
                              type: logoType,
                              transformation: { width: 76, height: 48 },
                            }),
                          ]
                        : [],
                  }),
                ],
              }),
              new TableCell({
                width: {
                  size: schoolTextWidth,
                  type: WidthType.DXA,
                },
                borders: noBorders,
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  centerLines(["Republic of the Philippines"], { size: 12 }),
                  centerLines(["DEPARTMENT OF EDUCATION"], {
                    bold: true,
                    size: 14,
                    color: "0038A8",
                  }),
                  centerLines(
                    [
                      profile?.region
                        ? `Region ${profile.region}`
                        : "Region XIII — CARAGA",
                    ],
                    { size: 12, color: "0038A8" },
                  ),
                  centerLines(
                    [profile?.school_name || "SCHOOL NAME"],
                    { bold: true, size: 13, color: "0038A8" },
                  ),
                ],
              }),
              // Empty right column balances the logo column so the heading
              // remains centered on the entire bond paper.
              new TableCell({
                width: {
                  size: schoolSideWidth,
                  type: WidthType.DXA,
                },
                borders: noBorders,
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [],
                  }),
                ],
              }),
            ],
          }),
        ],
      });

      const teacherColumnWidth = Math.floor(contentWidth * 0.75);
      const termColumnWidth = contentWidth - teacherColumnWidth;

      const teacherTermTable = new Table({
        width: { size: contentWidth, type: WidthType.DXA },
        columnWidths: [teacherColumnWidth, termColumnWidth],
        layout: TableLayoutType.FIXED,
        borders: noBorders,
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: teacherColumnWidth,
                  type: WidthType.DXA,
                },
                borders: noBorders,
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: "Teacher: ",
                        bold: true,
                        size: 12,
                        font: "Arial",
                      }),
                      new TextRun({
                        text: teacherName || "____________________",
                        color: "0038A8",
                        underline: {},
                        size: 12,
                        font: "Arial",
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: {
                  size: termColumnWidth,
                  type: WidthType.DXA,
                },
                borders: noBorders,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                      new TextRun({
                        text: "Term: ",
                        bold: true,
                        size: 12,
                        font: "Arial",
                      }),
                      new TextRun({
                        text: TERM_LABEL[term],
                        underline: {},
                        size: 12,
                        font: "Arial",
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      });

      const firstHeaderRow = new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          makeCell({
            text: ["Grade Level/", "Subjects"],
            width: columnWidths[0],
            bold: true,
            rowSpan: 2,
          }),
          makeCell({
            text: "STRAND",
            width: columnWidths[1],
            bold: true,
            rowSpan: 2,
          }),
          makeCell({
            text: "SECTION",
            width: columnWidths[2],
            bold: true,
            rowSpan: 2,
          }),
          makeCell({
            text: "REGISTERED LEARNERS",
            width: columnWidths.slice(3, 6).reduce((sum, value) => sum + value, 0),
            bold: true,
            columnSpan: 3,
          }),
          makeCell({
            text: ["LEARNERS WITH 75%", "AND ABOVE"],
            width: columnWidths.slice(6, 9).reduce((sum, value) => sum + value, 0),
            bold: true,
            columnSpan: 3,
          }),
          makeCell({
            text: "TOTAL GRADES",
            width: columnWidths.slice(9, 12).reduce((sum, value) => sum + value, 0),
            bold: true,
            columnSpan: 3,
          }),
          makeCell({
            text: "AVERAGE GRADE",
            width: columnWidths.slice(12, 15).reduce((sum, value) => sum + value, 0),
            bold: true,
            columnSpan: 3,
          }),
          makeCell({
            text: ["% OF PROFICIENCY", "(75%+)"],
            width: columnWidths.slice(15, 18).reduce((sum, value) => sum + value, 0),
            bold: true,
            columnSpan: 3,
          }),
        ],
      });

      const secondHeaderRow = new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: Array.from({ length: 5 }).flatMap((_, groupIndex) => {
          const start = 3 + groupIndex * 3;
          return [
            makeCell({
              text: "MALE",
              width: columnWidths[start],
              bold: true,
              size: 11,
            }),
            makeCell({
              text: "FEMALE",
              width: columnWidths[start + 1],
              bold: true,
              size: 11,
            }),
            makeCell({
              text: "TOTAL",
              width: columnWidths[start + 2],
              bold: true,
              size: 11,
            }),
          ];
        }),
      });

      const dataRows = rows.map((row) => {
        const totalRegistered = row.mReg + row.fReg;
        const totalPassed = row.mPass + row.fPass;
        const totalSum = row.mSum + row.fSum;
        const maleAverage = row.mReg ? row.mSum / row.mReg : 0;
        const femaleAverage = row.fReg ? row.fSum / row.fReg : 0;
        const totalAverage = totalRegistered ? totalSum / totalRegistered : 0;
        const maleProficiency = row.mReg
          ? Math.round((row.mPass / row.mReg) * 100)
          : 0;
        const femaleProficiency = row.fReg
          ? Math.round((row.fPass / row.fReg) * 100)
          : 0;
        const totalProficiency = totalRegistered
          ? Math.round((totalPassed / totalRegistered) * 100)
          : 0;

        const values = [
          row.subject,
          row.strand,
          row.section,
          String(row.mReg),
          String(row.fReg),
          String(totalRegistered),
          String(row.mPass),
          String(row.fPass),
          String(totalPassed),
          row.mSum.toFixed(0),
          row.fSum.toFixed(0),
          totalSum.toFixed(0),
          maleAverage ? maleAverage.toFixed(2) : "0",
          femaleAverage ? femaleAverage.toFixed(2) : "0",
          totalAverage ? totalAverage.toFixed(2) : "0",
          `${maleProficiency}%`,
          `${femaleProficiency}%`,
          `${totalProficiency}%`,
        ];

        return new TableRow({
          cantSplit: true,
          height: { value: 330, rule: HeightRule.ATLEAST },
          children: values.map((value, index) =>
            makeCell({
              text: value,
              width: columnWidths[index],
              bold: [5, 8, 11, 14, 17].includes(index),
              size: 11,
              alignment: index === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            }),
          ),
        });
      });

      const emptyRows = Array.from({
        length: Math.max(0, 3 - dataRows.length),
      }).map(
        () =>
          new TableRow({
            cantSplit: true,
            height: { value: 420, rule: HeightRule.ATLEAST },
            children: columnWidths.map((width) =>
              makeCell({ text: "", width, size: 11 }),
            ),
          }),
      );

      const gsaTable = new Table({
        width: { size: contentWidth, type: WidthType.DXA },
        columnWidths,
        layout: TableLayoutType.FIXED,
        rows: [firstHeaderRow, secondHeaderRow, ...dataRows, ...emptyRows],
      });

      const signatureColumnWidths = [
        Math.floor(contentWidth / 4),
        Math.floor(contentWidth / 4),
        Math.floor(contentWidth / 4),
        contentWidth - Math.floor(contentWidth / 4) * 3,
      ];

      const signatureCell = (
        label: string,
        name: string,
        title: string,
        width: number,
      ) =>
        new TableCell({
          width: {
            size: width,
            type: WidthType.DXA,
          },
          borders: noBorders,
          margins: { top: 0, bottom: 0, left: 80, right: 80 },
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({
                  text: label,
                  bold: true,
                  size: 11,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 260, after: 0 },
              border: {
                top: {
                  style: BorderStyle.SINGLE,
                  size: 6,
                  color: "000000",
                  space: 1,
                },
              },
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: name.toUpperCase(),
                  bold: true,
                  size: 12,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: title,
                  size: 10,
                  font: "Arial",
                }),
              ],
            }),
          ],
        });

      const signatureTable = new Table({
        width: { size: contentWidth, type: WidthType.DXA },
        columnWidths: signatureColumnWidths,
        layout: TableLayoutType.FIXED,
        borders: noBorders,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [
              signatureCell(
                "Prepared by:",
                signatories.prepared,
                signatories.preparedTitle,
                signatureColumnWidths[0],
              ),
              signatureCell(
                "Checked by:",
                signatories.checked,
                signatories.checkedTitle,
                signatureColumnWidths[1],
              ),
              signatureCell(
                "Noted:",
                signatories.noted,
                signatories.notedTitle,
                signatureColumnWidths[2],
              ),
              signatureCell(
                "Approved:",
                signatories.approved,
                signatories.approvedTitle,
                signatureColumnWidths[3],
              ),
            ],
          }),
        ],
      });

      const wordDocument = new WordDocument({
        styles: {
          default: {
            document: {
              run: { font: "Arial", size: 12, color: "000000" },
              paragraph: { spacing: { before: 0, after: 0, line: 240 } },
            },
          },
        },
        sections: [
          {
            properties: {
              page: {
                size: {
                  width: pageWidth,
                  height: pageHeight,
                  orientation: PageOrientation.LANDSCAPE,
                },
                margin: {
                  top: margin,
                  bottom: margin,
                  left: margin,
                  right: margin,
                },
              },
            },
            children: [
              schoolHeader,
              new Paragraph({
                spacing: { before: 60, after: 140 },
                border: {
                  bottom: {
                    style: BorderStyle.SINGLE,
                    size: 8,
                    color: "000000",
                    space: 1,
                  },
                },
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0, line: 240 },
                children: [
                  new TextRun({
                    text: "REPORT ON GENERAL SCHOLASTIC APTITUDE OF LEARNERS",
                    bold: true,
                    size: 15,
                    font: "Arial",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 30, line: 240 },
                children: [
                  new TextRun({
                    text: "PER TERM BY LEARNING AREA",
                    bold: true,
                    size: 15,
                    font: "Arial",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 140 },
                children: [
                  new TextRun({
                    text: `School Year: ${klass.school_year || "____-____"}`,
                    size: 11,
                    font: "Arial",
                  }),
                ],
              }),
              teacherTermTable,
              new Paragraph({ spacing: { after: 45 } }),
              gsaTable,
              new Paragraph({ spacing: { before: 220, after: 0 } }),
              signatureTable,
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(wordDocument);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `GSA_${klass.section || "class"}.docx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      console.error("Unable to create GSA Word document.", error);
      window.alert(
        error instanceof Error
          ? error.message
          : "Unable to create the GSA Word document.",
      );
    }
  };


  return (
    <div className="space-y-4 pb-8">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div className="flex items-center gap-3">
          <Link to="/school-forms" className="inline-flex items-center gap-1.5 rounded-lg border-2 px-3 py-1.5 text-sm font-medium text-white shadow-sm hover:opacity-90" style={{ backgroundColor: "#E11D48", borderColor: "#E11D48" }}>
            <ArrowLeft className="size-4" /> Back to Forms
          </Link>
          <div className="flex items-center gap-2 text-lg font-semibold">
            <FileText className="size-5 text-rose-600" /> GSA — General Scholastic Aptitude
          </div>
        </div>
      </div>


      <div
        className="relative overflow-hidden rounded-2xl border-2 px-5 py-3 shadow-sm sm:px-6 sm:py-0.5"
        style={{ backgroundColor: "#FFFBEB", borderColor: DEPED_YELLOW }}
      >
        <div className="text-sm font-semibold text-amber-900">
          Form Configuration
        </div>

        <div className="mx-auto mt-1 flex max-w-2xl flex-col items-center">
          <div className="flex w-full items-center justify-center gap-3 sm:gap-5">
            <div className="hidden h-px w-20 bg-amber-400 sm:block" />
            <div
              className="grid size-9 shrink-0 place-items-center rounded-full"
              style={{ backgroundColor: "#FEF0B6" }}
            >
              <GraduationCap className="size-5 text-amber-950" />
            </div>
            <div className="hidden h-px w-20 bg-amber-400 sm:block" />
          </div>

          <div className="mt-1 text-center text-xs font-bold uppercase tracking-[0.32em] text-amber-950">
            Class
          </div>

          <div className="mt-1 flex w-full max-w-[300px] items-center gap-1.5 rounded-lg border border-amber-300 bg-white/75 px-2 py-2 shadow-sm">
            <div className="grid size-5 shrink-0 place-items-center rounded-md bg-amber-50 text-amber-700">
              <Users className="size-3" />
            </div>

            <div className="min-w-0 flex-1 text-center">
              <div className="truncate text-[11px] font-bold text-amber-950">
                {klass
                  ? `${klass.grade_level || "—"} · ${klass.subject || "—"} · ${klass.section || "—"}`
                  : "No class selected"}
              </div>
            </div>
          </div>

          <p className="mt-0.5 text-center text-[9px] text-amber-900/60">
            Selected class for this form
          </p>
        </div>
      </div>

      <PdfPreviewShell
        fileName={`GSA_${klass?.section || "class"}.pdf`}
        docBaseName={`GSA_${klass?.section || "class"}`}
        printTargetId="gsa-doc"
        pageLabel={`${rows.length} subject${rows.length === 1 ? "" : "s"}`}
        paper={paper}
        onPaperChange={setPaper}
        onCopyToWord={copyGsaToWord}
      >
        <div id="gsa-doc" className="p-8 text-[10px] text-black" style={{ width: 1200, fontFamily: "Arial, Helvetica, sans-serif" }}>
          <DepEdHeader
            region={profile?.region ? `Region ${profile.region}` : "Region XIII — CARAGA"}
            division={profile?.division ? `Division of ${profile.division}` : ""}
            district={profile?.district || ""}
            schoolName={profile?.school_name || "SCHOOL NAME"}
          />
          <div className="mt-3 text-center text-[13px] font-bold">
            REPORT ON GENERAL SCHOLASTIC APTITUDE OF LEARNERS
            <div>PER TERM BY LEARNING AREA</div>
          </div>
          <div className="mt-1 text-center text-[10px]">School Year: {klass?.school_year || "____-____"}</div>

          <div className="mt-4 flex items-end justify-between text-[10px]">
            <div><b>Teacher:</b> <span className="ml-1 border-b border-black px-6" style={{ color: DEPED_BLUE }}>{teacherName}</span></div>
            <div><b>Term:</b> <u>{TERM_LABEL[term]}</u></div>
          </div>

          <table className="mt-2 w-full border border-black border-collapse text-[10px]">
            <thead>
              <tr>
                <th className="border border-black px-1 py-1 w-20" rowSpan={2}>Grade Level/<br />Subjects</th>
                <th className="border border-black px-1 py-1 w-16" rowSpan={2}>STRAND</th>
                <th className="border border-black px-1 py-1 w-16" rowSpan={2}>SECTION</th>
                <th className="border border-black px-1 py-1" colSpan={3}>REGISTERED LEARNERS</th>
                <th className="border border-black px-1 py-1" colSpan={3}>LEARNERS WITH 75% AND ABOVE</th>
                <th className="border border-black px-1 py-1" colSpan={3}>TOTAL GRADES</th>
                <th className="border border-black px-1 py-1" colSpan={3}>AVERAGE GRADE</th>
                <th className="border border-black px-1 py-1" colSpan={3}>% OF PROFICIENCY (75%+)</th>
              </tr>
              <tr>
                {Array.from({ length: 5 }).map((_, i) => (
                  <Fragment key={i}>
                    <th className="border border-black px-1 py-0.5 w-10">MALE</th>
                    <th className="border border-black px-1 py-0.5 w-10">FEMALE</th>
                    <th className="border border-black px-1 py-0.5 w-10">TOTAL</th>
                  </Fragment>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => {
                const tReg = r.mReg + r.fReg;
                const tPass = r.mPass + r.fPass;
                const tSum = r.mSum + r.fSum;
                const mAvg = r.mReg ? r.mSum / r.mReg : 0;
                const fAvg = r.fReg ? r.fSum / r.fReg : 0;
                const tAvg = tReg ? tSum / tReg : 0;
                const mProf = r.mReg ? Math.round((r.mPass / r.mReg) * 100) : 0;
                const fProf = r.fReg ? Math.round((r.fPass / r.fReg) * 100) : 0;
                const tProf = tReg ? Math.round((tPass / tReg) * 100) : 0;
                return (
                  <tr key={i}>
                    <td className="border border-black px-1 py-1">{r.subject}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.strand}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.section}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.mReg}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.fReg}</td>
                    <td className="border border-black px-1 py-1 text-center font-bold">{tReg}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.mPass}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.fPass}</td>
                    <td className="border border-black px-1 py-1 text-center font-bold">{tPass}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.mSum.toFixed(0)}</td>
                    <td className="border border-black px-1 py-1 text-center">{r.fSum.toFixed(0)}</td>
                    <td className="border border-black px-1 py-1 text-center font-bold">{tSum.toFixed(0)}</td>
                    <td className="border border-black px-1 py-1 text-center">{mAvg ? mAvg.toFixed(2) : 0}</td>
                    <td className="border border-black px-1 py-1 text-center">{fAvg ? fAvg.toFixed(2) : 0}</td>
                    <td className="border border-black px-1 py-1 text-center font-bold">{tAvg ? tAvg.toFixed(2) : 0}</td>
                    <td className="border border-black px-1 py-1 text-center">{mProf}%</td>
                    <td className="border border-black px-1 py-1 text-center">{fProf}%</td>
                    <td className="border border-black px-1 py-1 text-center font-bold">{tProf}%</td>
                  </tr>
                );
              })}
              {Array.from({ length: Math.max(0, 3 - rows.length) }).map((_, i) => (
                <tr key={`e${i}`}>{Array.from({ length: 18 }).map((_, j) => <td key={j} className="border border-black px-1 py-3">&nbsp;</td>)}</tr>
              ))}
            </tbody>
          </table>

          <div className="mt-10 grid grid-cols-4 gap-6 text-center text-[10px]">
            {([
              ["Prepared by:", signatories.prepared, signatories.preparedTitle],
              ["Checked by:", signatories.checked, signatories.checkedTitle],
              ["Noted:", signatories.noted, signatories.notedTitle],
              ["Approved:", signatories.approved, signatories.approvedTitle],
            ] as const).map(([label, name, title]) => (
              <div key={label}>
                <div className="text-left font-semibold">{label}</div>
                <div className="mt-6 border-t border-black pt-1 text-[11px] font-bold uppercase">{name}</div>
                <div className="text-[9px]">{title}</div>
              </div>
            ))}
          </div>
        </div>
      </PdfPreviewShell>
    </div>
  );
}
