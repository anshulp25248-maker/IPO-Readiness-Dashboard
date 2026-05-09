import Link from "next/link";

type PlaceholderPageProps = {
  title: string;
};

export function PlaceholderPage({ title }: PlaceholderPageProps) {
  return (
    <main className="min-h-screen bg-[linear-gradient(135deg,#7dd3fc,#99f6e4,#dcfce7)] px-6 py-8 text-slate-950">
      <Link
        href="/"
        className="inline-flex h-11 items-center rounded-lg border border-white/40 bg-white/25 px-5 text-sm font-semibold shadow-lg backdrop-blur-xl transition hover:-translate-y-0.5 hover:bg-white/35"
      >
        Back
      </Link>
      <section className="mx-auto flex min-h-[70vh] max-w-4xl flex-col items-center justify-center text-center">
        <h1 className="font-serif text-5xl font-bold text-slate-950 sm:text-6xl">
          {title}
        </h1>
      </section>
    </main>
  );
}
