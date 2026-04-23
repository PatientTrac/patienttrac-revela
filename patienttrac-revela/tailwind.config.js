/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        revela: {
          navy:    '#0a1628',
          slate:   '#1a3a5c',
          gold:    '#c9a96e',
          ivory:   '#f0ece4',
          teal:    '#4a7fa8',
          'navy-light': '#142035',
          'gold-light': '#fdf9f0',
          'teal-light': '#e8f2f8',
        },
      },
      fontFamily: {
        serif:  ['Georgia', 'Cambria', 'Times New Roman', 'serif'],
        sans:   ['Inter', 'system-ui', 'sans-serif'],
        mono:   ['JetBrains Mono', 'Fira Code', 'monospace'],
      },
      borderRadius: {
        'revela': '10px',
      },
    },
  },
  plugins: [],
}
