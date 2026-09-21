import depedLogo from "@/assets/deped.jpg";

export const DEPED_BLUE = "#0038A8";
export const DEPED_RED = "#CE1126";
export const DEPED_YELLOW = "#FCD116";

type Props = {
  region?: string | null;
  division?: string | null;
  district?: string | null;
  schoolName?: string | null;
  title?: string;
  subtitle?: string;
  compact?: boolean;
};

export function DepEdHeader({
  region, division, district, schoolName, title, subtitle, compact,
}: Props) {
  return (
    <div className="text-center" style={{ color: "#111" }}>
      <div className="flex items-center justify-center gap-3"style={{ transform: "translateX(-65px)" }}>
        <img src={depedLogo} alt="DepEd" className={compact ? "h-10" : "h-14"} />
        <div>
          <div className="text-[10px] tracking-wide">Republic of the Philippines</div>
          <div className="text-[13px] font-bold" style={{ color: DEPED_BLUE }}>DEPARTMENT OF EDUCATION</div>
          {region && <div className="text-[10px]">{region}</div>}
          {division && <div className="text-[10px]">{division}</div>}
          {district && <div className="text-[10px]">{district}</div>}
          {schoolName && (
            <div className="text-[12px] font-bold uppercase" style={{ color: DEPED_BLUE }}>
              {schoolName}
            </div>
          )}
        </div>
      </div>
      {title && (
        <div className="mt-2 text-[13px] font-bold uppercase tracking-wide" style={{ color: DEPED_BLUE }}>
          {title}
        </div>
      )}
      {subtitle && <div className="text-[10px] italic">{subtitle}</div>}
    </div>
  );
}
