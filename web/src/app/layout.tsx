import type { Metadata } from "next";
import { Inter, JetBrains_Mono } from "next/font/google";
import AuthGate from "@/components/AuthGate";
import "./globals.css";

const inter = Inter({
  variable: "--font-inter",
  subsets: ["latin"],
});

const jetbrainsMono = JetBrains_Mono({
  variable: "--font-jetbrains",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "United Glass Ventures | Estimator",
  description: "Professional Storefront & Curtain Wall Cost Estimation Tool",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark">
      <body
        className={`${inter.variable} ${jetbrainsMono.variable} antialiased min-h-screen`}
      >
        <AuthGate>{children}</AuthGate>
      </body>
    </html>
  );
}
