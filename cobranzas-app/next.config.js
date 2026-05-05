/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  env: {
    GOOGLE_SHEET_URL: process.env.GOOGLE_SHEET_URL,
  },
};

module.exports = nextConfig;
