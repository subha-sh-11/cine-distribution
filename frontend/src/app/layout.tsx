import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "SVF Distribution — Admin",
  description: "Nizam territory theatre & distribution admin",
};

const themeScript = `
(function(){
  try {
    var t = localStorage.getItem('svf.theme') || 'light';
    document.documentElement.setAttribute('data-theme', t);
    document.documentElement.setAttribute('data-accent', 'amber');
  } catch (e) {}               
})();
`;

export default function RootLayout({
  children,
}: {          
  children: React.ReactNode;
}) {       
  return (
    <html lang="en" data-theme="light" data-accent="amber">
      <head>
        <script dangerouslySetInnerHTML={{ __html: themeScript }} />
      </head>
      <body>{children}</body>
    </html>
  );
}
