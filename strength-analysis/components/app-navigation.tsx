"use client";

import Link from "next/link";
import Image from "next/image";
import { usePathname } from "next/navigation";
import { cn } from "@/lib/utils";

export function AppNavigation() {
  const path = usePathname();
  return <div className="border-b border-neutral-200 bg-white"><div className="mx-auto flex max-w-[1440px] flex-wrap items-center justify-between gap-4 px-5 py-3 sm:px-8 lg:px-10">
    <Link href="/" aria-label="Kohinoor Textile Mills home" className="flex items-center gap-4"><Image src="/kohinoor-logo.png" alt="Kohinoor Textile Mills Limited" width={224} height={52} priority unoptimized className="h-13 w-56 shrink-0" /><span className="hidden border-l border-neutral-200 pl-4 text-sm font-medium text-neutral-600 sm:block">Workforce reporting</span></Link>
    <nav aria-label="Main navigation" className="flex gap-1">{[{ href: "/", label: "Turnover report" }, { href: "/config", label: "Configuration" }].map((item) => <Link key={item.href} href={item.href} aria-current={path === item.href ? "page" : undefined} className={cn("rounded-md px-3 py-2 text-sm transition-colors", path === item.href ? "bg-neutral-100 font-medium text-neutral-950" : "text-neutral-500 hover:bg-neutral-50")}>{item.label}</Link>)}</nav>
  </div></div>;
}
