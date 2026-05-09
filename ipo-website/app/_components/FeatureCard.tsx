import Link from "next/link";

type FeatureCardProps = {
  title: string;
  href: string;
};

export function FeatureCard({ title, href }: FeatureCardProps) {
  return (
    <Link
      href={href}
      className="group flex min-h-24 w-full items-center justify-center rounded-2xl border border-white/35 bg-white/15 px-5 py-6 text-center shadow-[0_20px_60px_rgba(15,23,42,0.18)] backdrop-blur-xl transition duration-300 ease-out hover:-translate-y-1 hover:scale-[1.025] hover:border-white/70 hover:shadow-[0_24px_80px_rgba(30,64,175,0.28)] focus:outline-none focus:ring-4 focus:ring-white/45 sm:min-h-32 sm:px-6 sm:py-8"
    >
      <span className="text-xl font-bold tracking-wide text-slate-950 transition duration-300 group-hover:text-indigo-950 sm:text-2xl">
        {title}
      </span>
    </Link>
  );
}
