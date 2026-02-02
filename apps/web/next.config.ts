import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  async rewrites() {
    const raw = process.env.API_ORIGIN || "http://localhost:3001";
    const apiOrigin =
      raw.startsWith("http://") || raw.startsWith("https://")
        ? raw
        : `http://${raw}`;
    return [
      {
        source: "/api/:path*",
        destination: `${apiOrigin}/api/:path*`,
      },
    ];
  },
};

export default nextConfig;
