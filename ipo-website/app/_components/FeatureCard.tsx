import Link from "next/link";

type FeatureCardProps = {
  title: string;
  href: string;
  icon: "dashboard" | "report" | "ranking" | "profile" | "competitors";
  color: string;
  soft: string;
};

function FeatureIcon({ icon }: { icon: FeatureCardProps["icon"] }) {
  const iconClass = "h-8 w-8 sm:h-10 sm:w-10 lg:h-12 lg:w-12";

  if (icon === "dashboard") {
    return (
      <svg className={iconClass} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M4 13h6V4H4v9Z" />
        <path d="M14 20h6V4h-6v16Z" />
        <path d="M4 20h6v-3H4v3Z" />
      </svg>
    );
  }

  if (icon === "report") {
    return (
      <svg className={iconClass} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M7 3h7l4 4v14H7V3Z" />
        <path d="M14 3v5h5" />
        <path d="M10 12h6" />
        <path d="M10 16h4" />
      </svg>
    );
  }

  if (icon === "ranking") {
    return (
      <svg className={iconClass} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M5 20v-6h4v6H5Z" />
        <path d="M10 20V4h4v16h-4Z" />
        <path d="M15 20v-9h4v9h-4Z" />
      </svg>
    );
  }

  if (icon === "profile") {
    return (
      <svg className={iconClass} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="M12 12a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z" />
        <path d="M4 21a8 8 0 0 1 16 0" />
      </svg>
    );
  }

  return (
    <svg className={iconClass} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
      <path d="M8 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8Z" />
      <path d="M16 13a3 3 0 1 0 0-6 3 3 0 0 0 0 6Z" />
      <path d="M2 21a6 6 0 0 1 12 0" />
      <path d="M14 20a5 5 0 0 1 8 0" />
    </svg>
  );
}

export function FeatureCard({ title, href, icon, color, soft }: FeatureCardProps) {
  return (
    <Link
      href={href}
      className="group relative flex min-h-24 w-full flex-col items-center justify-center gap-2 overflow-hidden rounded-2xl border px-4 py-4 text-center shadow-[0_16px_44px_rgba(15,23,42,0.12)] transition duration-300 ease-out hover:-translate-y-1 hover:scale-[1.025] hover:shadow-[0_24px_80px_rgba(15,23,42,0.18)] focus:outline-none focus:ring-4 focus:ring-slate-200 sm:min-h-28 sm:gap-3 sm:px-5 sm:py-5 lg:min-h-32"
      style={{ backgroundColor: soft, borderColor: color, color }}
    >
      <span className="absolute left-0 top-0 h-1.5 w-full transition duration-300 group-hover:h-2.5" style={{ backgroundColor: color }} />
      <FeatureIcon icon={icon} />
      <span className="text-base font-black tracking-wide transition duration-300 sm:text-lg lg:text-xl">{title}</span>
    </Link>
  );
}
