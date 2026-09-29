/**
 * Тема: светлая / тёмная.
 * Сохраняется в localStorage, применяется при загрузке.
 *
 * Поддерживает кнопки: #themeToggle, #themeToggleTop, #themeToggleSidebar.
 * Синхронно ставит и класс .dark, и data-theme="dark" на <html>.
 */
(function () {
    const STORAGE_KEY = 'theme';
    const DEFAULT = 'light';

    function applyTheme(theme) {
        const html = document.documentElement;

        if (theme === 'dark') {
            html.classList.add('dark');
            html.setAttribute('data-theme', 'dark');
        } else {
            html.classList.remove('dark');
            html.setAttribute('data-theme', 'light');
        }

        // Обновляем эмодзи на всех кнопках переключения
        ['themeToggle', 'themeToggleTop', 'themeToggleSidebar'].forEach(id => {
            const btn = document.getElementById(id);
            if (!btn) return;
            // Если внутри кнопки два <span> — не трогаем, там CSS-логика (.dark:hidden / .dark:block)
            // Если кнопка пустая или текстом — ставим эмодзи
            const hasSpans = btn.querySelector('span');
            if (!hasSpans) {
                btn.textContent = theme === 'dark' ? '☀️' : '🌙';
            }
        });
    }

    function getSavedTheme() {
        const saved = localStorage.getItem(STORAGE_KEY);
        if (saved === 'dark' || saved === 'light') return saved;
        // Если пользователь не выбирал — используем системную
        if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
            return 'dark';
        }
        return DEFAULT;
    }

    function toggleTheme() {
        const current = document.documentElement.classList.contains('dark') ? 'dark' : 'light';
        const next = current === 'dark' ? 'light' : 'dark';
        localStorage.setItem(STORAGE_KEY, next);
        applyTheme(next);
    }

    // Применяем до отрисовки
    applyTheme(getSavedTheme());

    // Вешаем обработчики после DOMContentLoaded
    document.addEventListener('DOMContentLoaded', function () {
        applyTheme(getSavedTheme());

        ['themeToggle', 'themeToggleTop', 'themeToggleSidebar'].forEach(id => {
            const btn = document.getElementById(id);
            if (btn) btn.addEventListener('click', toggleTheme);
        });
    });

    // Публичный API на случай, если нужно вызвать извне
    window.__theme = { applyTheme, toggleTheme, getSavedTheme };
})();