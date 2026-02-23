import type { Metadata } from "next";
import { Inter, JetBrains_Mono } from "next/font/google";
import AuthGate from "@/components/AuthGate";
import { BubbleBackground } from "@/components/BubbleBackground";
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
        className={`${inter.variable} ${jetbrainsMono.variable} antialiased min-h-screen bg-[#06060a]`}
      >
        <BubbleBackground
          interactive
          className="fixed inset-0 z-0 bg-gradient-to-br from-[#03030a] to-[#06061a]"
          colors={{
            first: '10,40,100',
            second: '50,30,110',
            third: '5,70,70',
            fourth: '20,25,70',
            fifth: '40,25,100',
            sixth: '30,65,130',
          }}
        />
        <div className="fixed inset-0 z-[1] bg-[#06060a]/60 pointer-events-none" />
        <div className="relative z-10 min-h-screen">
          <AuthGate>{children}</AuthGate>
        </div>
      </body>
    </html>
  );
}
