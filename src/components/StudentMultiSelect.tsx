import { useMemo, useState } from "react";
import { Search, Check, ChevronDown, X } from "lucide-react";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import type { StudentRow } from "@/lib/data";

type Props = {
  students: StudentRow[];
  selected: Set<string>;
  onChange: (next: Set<string>) => void;
  color?: string;
  label?: string;
  className?: string;
};

export function StudentMultiSelect({
  students, selected, onChange, color = "#0038A8", label = "Learners", className,
}: Props) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState("");

  const filtered = useMemo(() => {
    const t = q.trim().toLowerCase();
    if (!t) return students;
    return students.filter((s) =>
      `${s.first_name} ${s.middle_name || ""} ${s.last_name} ${s.lrn || ""}`.toLowerCase().includes(t),
    );
  }, [students, q]);

  const toggle = (id: string) => {
    const n = new Set(selected);
    n.has(id) ? n.delete(id) : n.add(id);
    onChange(n);
  };

  const selectedList = students.filter((s) => selected.has(s.id));
  const summary =
    selected.size === 0
      ? `Select ${label.toLowerCase()}…`
      : selected.size === students.length
      ? `All ${students.length} ${label.toLowerCase()}`
      : `${selected.size} of ${students.length} selected`;

  return (
    <div className={className}>
      <Popover open={open} onOpenChange={setOpen}>
        <PopoverTrigger asChild>
          <button
            type="button"
            className="flex w-full items-center justify-between gap-2 rounded-lg border-2 bg-white px-3 py-2 text-left text-sm hover:bg-slate-50"
            style={{ borderColor: color }}
          >
            <span className="truncate font-medium" style={{ color }}>{summary}</span>
            <ChevronDown className="size-4 shrink-0" style={{ color }} />
          </button>
        </PopoverTrigger>
        <PopoverContent align="start" className="w-80 p-0 pointer-events-auto">
          <div className="flex items-center gap-2 border-b px-3 py-2">
            <Search className="size-4 text-slate-400" />
            <input
              autoFocus
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="Search by name or LRN…"
              className="flex-1 bg-transparent text-sm outline-none placeholder:text-slate-400"
            />
          </div>
          <div className="flex items-center justify-between gap-2 border-b px-3 py-1.5">
            <button
              type="button"
              onClick={() => onChange(new Set(students.map((s) => s.id)))}
              className="rounded-md border-2 px-2 py-0.5 text-[11px] font-semibold"
              style={{ borderColor: color, color, backgroundColor: `${color}12` }}
            >
              Select All
            </button>
            <button
              type="button"
              onClick={() => onChange(new Set())}
              className="rounded-md border px-2 py-0.5 text-[11px] font-medium hover:bg-muted"
            >
              Clear
            </button>
            <span className="ml-auto text-[10px] text-slate-500">{filtered.length} shown</span>
          </div>
          <div className="max-h-64 overflow-auto py-1">
            {filtered.length === 0 && (
              <div className="px-3 py-4 text-center text-xs text-slate-400">No matching students.</div>
            )}
            {filtered.map((s) => {
              const checked = selected.has(s.id);
              return (
                <button
                  key={s.id}
                  type="button"
                  onClick={() => toggle(s.id)}
                  className="flex w-full items-center gap-2 px-3 py-1.5 text-left text-sm hover:bg-slate-50"
                >
                  <span
                    className="grid size-4 place-items-center rounded border"
                    style={{
                      borderColor: checked ? color : "#cbd5e1",
                      backgroundColor: checked ? color : "transparent",
                    }}
                  >
                    {checked && <Check className="size-3 text-white" />}
                  </span>
                  <span className="truncate">{s.last_name}, {s.first_name} {s.middle_name || ""}</span>
                </button>
              );
            })}
          </div>
        </PopoverContent>
      </Popover>

      {selectedList.length > 0 && (
        <div className="mt-2 flex flex-wrap gap-1.5">
          {selectedList.map((s) => (
            <span
              key={s.id}
              className="inline-flex items-center gap-1 rounded-full border px-2 py-0.5 text-[11px]"
              style={{ borderColor: color, color, backgroundColor: `${color}10` }}
            >
              {s.first_name} {s.last_name}
              <button
                type="button"
                onClick={() => toggle(s.id)}
                className="rounded-full hover:bg-black/10"
                aria-label="Remove"
              >
                <X className="size-3" />
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
