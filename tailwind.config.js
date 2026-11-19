import defaultTheme from 'tailwindcss/defaultTheme';

/** @type {import('tailwindcss').Config} */
// Scans index.html + all source JS for class names. Note: the app builds markup as template
// strings, so class names must appear as literal strings in source to be detected (they do).
export default {
  content: ['./index.html', './src/**/*.{js,ts}'],
  theme: {
    extend: {
      // Sarabun for Thai and Latin (user-approved 19 Nov 2026); the system fonts stay as the fallback
      fontFamily: { sans: ['Sarabun', ...defaultTheme.fontFamily.sans] },
      colors: {
        'primary': '#0284c7',   // sky-600
        'secondary': '#0369a1', // sky-700
        'accent': '#f97316',    // orange-500
        'neutral': '#374151',   // gray-700
        'base-100': '#1f2937',  // gray-800
        'info': '#0ea5e9',      // sky-500
        'success': '#22c55e',   // green-500
        'warning': '#eab308',   // yellow-500
        'error': '#ef4444',     // red-500
      },
    },
  },
  plugins: [],
};
