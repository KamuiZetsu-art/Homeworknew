/**
 * Общие функции для всего приложения.
 *
 * Включает:
 *  - Toast-уведомления (showToast)
 *  - Модальные окна (openModal/closeModal) — совместимы с обоими стилями:
 *      * class="show"     (design-system.css)
 *      * style="display:flex" (inline в шаблонах)
 *  - Бургер-меню для мобильных
 *  - Утилиты: escapeHtml, escapeAttr, confirmDelete
 */

// ==================== TOAST ====================
function showToast(message, type = 'info') {
    let toast = document.getElementById('toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toast';
        document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = 'toast ' + type;
    toast.style.display = 'block';

    clearTimeout(toast._hideTimeout);
    toast._hideTimeout = setTimeout(() => {
        toast.style.display = 'none';
    }, 3000);
}

// ==================== MODAL ====================
// Универсально: работает и с .show, и с display:flex / display:none
function openModal(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.add('show');
    // Если шаблон использует inline display — ставим flex
    el.style.display = 'flex';
}

function closeModal(id) {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.remove('show');
    el.style.display = 'none';
}

// Закрытие по клику на backdrop
document.addEventListener('click', function (event) {
    const t = event.target;
    if (t && t.classList && t.classList.contains('modal')) {
        t.classList.remove('show');
        t.style.display = 'none';
    }
});

// Закрытие по Escape
document.addEventListener('keydown', function (event) {
    if (event.key !== 'Escape') return;
    document.querySelectorAll('.modal.show').forEach(m => {
        m.classList.remove('show');
        m.style.display = 'none';
    });
    // На случай, если .show не было, но display:flex осталось
    document.querySelectorAll('.modal').forEach(m => {
        if (m.style.display === 'flex') m.style.display = 'none';
    });
});

// ==================== БУРГЕР-МЕНЮ ====================
function toggleMobileNav() {
    const nav = document.getElementById('mainNav');
    if (!nav) return;
    nav.classList.toggle('mobile-open');
}

document.addEventListener('DOMContentLoaded', function () {
    const nav = document.getElementById('mainNav');
    if (nav) {
        nav.querySelectorAll('a').forEach(link => {
            link.addEventListener('click', () => {
                nav.classList.remove('mobile-open');
            });
        });
    }
    // Автоскрытие прелоадеров
    document.querySelectorAll('[data-auto-hide]').forEach(el => {
        setTimeout(() => { el.style.display = 'none'; }, 500);
    });
});

// ==================== УТИЛИТЫ ====================
function confirmDelete(message) {
    return confirm(message || 'Удалить?');
}

/** Экранирование текста для вставки в HTML (защита от XSS). */
function escapeHtml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/** Экранирование для вставки в атрибут (в т.ч. в inline onclick). */
function escapeAttr(value) {
    return escapeHtml(value);
}

/** Безопасная вставка в JSON-контекст скрипта. */
function jsString(value) {
    if (value === null || value === undefined) return '""';
    return JSON.stringify(String(value));
}