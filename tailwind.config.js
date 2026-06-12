/** @type {import('tailwindcss').Config} */
// 旧 index.html 内のインライン設定（CDN版）から移行したビルド設定。
// クラスは index.html と assets/js/*.js（テンプレート文字列）から抽出される。
module.exports = {
    darkMode: 'class',
    content: ['./index.html', './assets/js/**/*.js'],
    // displayAIAdvice() が `bg-${statusColor}-50` 形式で動的生成するクラス
    safelist: [
        ...['green', 'blue', 'yellow', 'red'].flatMap(c => [
            `bg-${c}-50`, `border-${c}-500`, `text-${c}-800`, `text-${c}-600`,
        ]),
    ],
    theme: {
        extend: {
            colors: {
                // Elegant Salon Theme - Warm & Sophisticated
                primary: {
                    50: '#faf8f5',
                    100: '#f5efe8',
                    200: '#ebe0d1',
                    300: '#dcc9b3',
                    400: '#c9a87e',
                    500: '#b8956a',  // Champagne Gold
                    600: '#a07d52',
                    700: '#866644',
                    800: '#6d533a',
                    900: '#5a4532',
                },
                accent: {
                    50: '#f5f6f8',
                    100: '#e8ebf0',
                    200: '#d4dae3',
                    300: '#b5c0cf',
                    400: '#8f9fb5',
                    500: '#6e819c',
                    600: '#566882',
                    700: '#47566b',  // Slate Blue
                    800: '#3d4859',
                    900: '#353e4c',
                },
                surface: {
                    50: '#faf9f7',
                    100: '#f6f4f1',
                    200: '#edeae5',
                    300: '#dedad3',
                    400: '#c7c1b6',
                    500: '#aea69a',
                    600: '#948a7d',
                    700: '#7a7167',
                    800: '#655d55',
                    900: '#544e47',
                },
                sage: {
                    400: '#8ba88e',
                    500: '#739977',  // Sage Green
                    600: '#5d7d60',
                },
                rose: {
                    400: '#c4a5a0',
                    500: '#b08f8a',  // Dusty Rose
                    600: '#9a7873',
                },
                warmgold: {
                    400: '#d4b896',
                    500: '#c9a96e',  // Warm Gold
                    600: '#b8941d',
                },
            },
            fontFamily: {
                display: ['"Cormorant Garamond"', 'Georgia', 'serif'],
                sans: ['"Inter"', '-apple-system', 'BlinkMacSystemFont', 'sans-serif'],
            },
            boxShadow: {
                'soft': '0 2px 15px -3px rgba(0, 0, 0, 0.07), 0 10px 20px -2px rgba(0, 0, 0, 0.04)',
                'soft-lg': '0 10px 40px -10px rgba(0, 0, 0, 0.1), 0 2px 10px -2px rgba(0, 0, 0, 0.04)',
                'glow': '0 0 20px rgba(184, 149, 106, 0.15)',
                'glow-accent': '0 0 20px rgba(86, 104, 130, 0.2)',
            }
        }
    }
};
