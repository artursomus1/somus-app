/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        somus: {
          // Primary brand - darker, more serious greens
          green: {
            50: '#E8F5EE',
            100: '#C8E6D5',
            200: '#93CDAB',
            300: '#5EB381',
            400: '#369A5D',
            500: '#1A7A3E',
            600: '#0D5C2C',
            700: '#003D1E',
            800: '#002E16',
            900: '#001F0E',
            950: '#001409',
            DEFAULT: '#1A7A3E',
            light: '#369A5D',
            lighter: '#93CDAB',
          },
          // Background system (dark mode financial BI)
          bg: {
            primary: '#0A0F14',
            secondary: '#0F1419',
            tertiary: '#151C24',
            input: '#1A2332',
            hover: '#1E2A3A',
          },
          // Accent colors from NASA spreadsheet
          purple: '#7030A0',
          navy: '#002060',
          teal: '#1B6B5F',
          gold: '#D4A017',
          orange: '#ED7D31',
          skyblue: '#00B0F0',
          red: '#C00000',
          // Text
          text: {
            primary: '#E8ECF0',
            secondary: '#8B95A5',
            tertiary: '#5A6577',
            accent: '#4ADE80',
            warning: '#F59E0B',
            danger: '#EF4444',
          },
          // Borders
          border: {
            DEFAULT: '#1E2A3A',
            light: '#2A3544',
            strong: '#3A4A5A',
          },
          // Legacy compat
          dark: '#0A0F14',
          gray: {
            50: '#151C24',
            100: '#1A2332',
            200: '#1E2A3A',
            300: '#2A3544',
            400: '#5A6577',
            500: '#8B95A5',
            600: '#A0A8B5',
            700: '#C0C6D0',
            800: '#E0E4EA',
            900: '#E8ECF0',
          },
        },
      },
      fontFamily: {
        sans: ['DM Sans', 'Inter', 'system-ui', 'sans-serif'],
        mono: ['JetBrains Mono', 'Fira Code', 'monospace'],
      },
      boxShadow: {
        'glow-green': '0 0 20px rgba(26, 122, 62, 0.15)',
        'glow-purple': '0 0 20px rgba(112, 48, 160, 0.15)',
        'glow-gold': '0 0 20px rgba(212, 160, 23, 0.15)',
      },
      backgroundImage: {
        'gradient-green': 'linear-gradient(135deg, #1A7A3E 0%, #0D5C2C 100%)',
        'gradient-green-subtle': 'linear-gradient(135deg, rgba(26,122,62,0.15) 0%, rgba(13,92,44,0.05) 100%)',
        'gradient-purple': 'linear-gradient(135deg, #7030A0 0%, #501878 100%)',
        'gradient-navy': 'linear-gradient(135deg, #002060 0%, #001540 100%)',
      },
    },
  },
  plugins: [],
};
