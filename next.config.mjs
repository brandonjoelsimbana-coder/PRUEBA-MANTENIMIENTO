/** @type {import('next').NextConfig} */
const nextConfig = {
  outputFileTracingIncludes: {
    "/api/ask": ["./data/**/*.xlsx"]
  }
};

export default nextConfig;
