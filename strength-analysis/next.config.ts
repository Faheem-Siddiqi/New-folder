import type { NextConfig } from "next";
import { resolve } from "node:path";

const nextConfig: NextConfig = {
  // Include the shared root template when tracing server files for deployment.
  outputFileTracingRoot: resolve(process.cwd(), ".."),
  outputFileTracingIncludes: { "/*": ["../strength.json", "../strength-last-result.json"] },
  experimental: { serverActions: { bodySizeLimit: "64mb" } },
  turbopack: { root: resolve(process.cwd(), "..") },
};

export default nextConfig;
