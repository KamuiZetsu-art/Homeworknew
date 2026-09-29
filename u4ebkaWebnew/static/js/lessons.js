function escapeHtml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function renderLessonCard(paraNum, lesson) {
    const subject = escapeHtml(lesson.предмет || '');
    const teacher = escapeHtml(lesson.преподаватель_коротко || lesson.преподаватель || '');
    const room = escapeHtml(lesson.кабинет || '');
    const time = escapeHtml(lesson.время || '');
    const comment = lesson.override && lesson.override.comment ? escapeHtml(lesson.override.comment) : '';

    if (lesson.отменено) {
        return `
        <div class="flex items-start gap-3 p-4 rounded-2xl bg-red-500/10 border border-red-500/20 border-l-4 border-l-red-500 shadow-sm">
            <div class="w-10 h-10 rounded-xl bg-red-500 text-white flex items-center justify-center font-bold flex-shrink-0 shadow-md">
                ${paraNum}
            </div>
            <div class="flex-1 min-w-0">
                <div class="font-bold line-through text-red-600 dark:text-red-400 text-sm">${subject}</div>
                <div class="text-xs mt-1 font-semibold text-red-500">Пара отменена</div>
                ${comment ? `<div class="text-[11px] mt-1 text-red-500 opacity-80">${comment}</div>` : ''}
            </div>
            <span class="px-2 py-1 rounded-lg bg-red-500/20 text-red-700 dark:text-red-300 text-[10px] font-bold">${time}</span>
        </div>`;
    }

    if (lesson.override && lesson.override.type === 'replaced') {
        return `
        <div class="flex items-start gap-3 p-4 rounded-2xl bg-orange-500/10 border border-orange-500/20 border-l-4 border-l-orange-500 shadow-sm">
            <div class="w-10 h-10 rounded-xl bg-orange-500 text-white flex items-center justify-center font-bold flex-shrink-0 shadow-md">
                ${paraNum}
            </div>
            <div class="flex-1 min-w-0">
                <div class="flex items-center gap-2 mb-1 flex-wrap">
                    <div class="font-bold text-slate-900 dark:text-white text-sm">${subject}</div>
                    <span class="px-2 py-0.5 rounded border border-orange-500/30 bg-orange-500/20 text-orange-600 dark:text-orange-400 text-[9px] font-black uppercase">Замена</span>
                </div>
                <div class="text-xs text-slate-500 dark:text-slate-400">
                    ${teacher} &nbsp;•&nbsp; ${room}
                </div>
                ${comment ? `<div class="text-[11px] mt-1 text-orange-600 dark:text-orange-400">${comment}</div>` : ''}
            </div>
            <span class="px-2 py-1 rounded-lg bg-slate-200/50 dark:bg-slate-700/50 text-slate-700 dark:text-slate-300 text-[10px] font-bold">${time}</span>
        </div>`;
    }

    return `
    <div class="flex items-start gap-3 p-4 rounded-2xl bg-white/50 dark:bg-slate-900/50 border border-white/60 dark:border-slate-800 border-l-4 border-l-cyan-500 shadow-sm">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-cyan-400 to-cyan-600 text-white flex items-center justify-center font-bold flex-shrink-0 shadow-md">
            ${paraNum}
        </div>
        <div class="flex-1 min-w-0">
            <div class="font-bold text-slate-900 dark:text-white text-sm">${subject}</div>
            <div class="text-xs text-slate-500 dark:text-slate-400 mt-1">
                ${teacher} &nbsp;•&nbsp; ${room}
            </div>
        </div>
        <span class="px-2 py-1 rounded-lg bg-cyan-500/10 text-cyan-600 dark:text-cyan-400 border border-cyan-500/20 text-[10px] font-bold">${time}</span>
    </div>`;
}

function renderLessonCardSmall(lesson) {
    const subject = escapeHtml(lesson.предмет || '');
    const teacher = escapeHtml(lesson.преподаватель_коротко || lesson.преподаватель || '');
    const room = escapeHtml(lesson.кабинет || '');
    const time = escapeHtml(lesson.время || '');

    if (lesson.отменено) {
        return `<div class="p-3 rounded-xl bg-red-500/10 border border-red-500/30 border-l-4 border-l-red-500 h-full flex flex-col justify-center text-[10px] shadow-sm">
            <div class="font-bold line-through text-red-600 dark:text-red-400 mb-1 leading-tight">${subject}</div>
            <div class="font-bold text-red-500 uppercase">Отменено</div>
            <div class="font-bold text-red-400 mt-1">${time}</div>
        </div>`;
    }
    if (lesson.override && lesson.override.type === 'replaced') {
        return `<div class="p-3 rounded-xl bg-orange-500/10 border border-orange-500/30 border-l-4 border-l-orange-500 h-full flex flex-col justify-center text-[10px] shadow-sm">
            <div class="font-bold text-orange-600 dark:text-orange-400 mb-1 leading-tight">${subject}</div>
            <div class="text-slate-600 dark:text-slate-400">${teacher}</div>
            <div class="text-slate-500">${room}</div>
            <div class="font-bold text-emerald-600 dark:text-emerald-400 mt-1">${time}</div>
        </div>`;
    }
    return `<div class="p-3 rounded-xl bg-white/60 dark:bg-slate-900/60 border border-white/80 dark:border-slate-800 border-l-4 border-l-cyan-500 h-full flex flex-col justify-center text-[10px] shadow-sm hover:shadow-md transition-all">
        <div class="font-bold text-slate-900 dark:text-white mb-1 leading-tight">${subject}</div>
        <div class="text-slate-500 dark:text-slate-400">${teacher}</div>
        <div class="text-slate-400">${room}</div>
        <div class="font-bold text-emerald-600 dark:text-emerald-400 mt-1">${time}</div>
    </div>`;
}

function renderLessonCardMobile(paraNum, lesson) {
    const subject = escapeHtml(lesson.предмет || '');
    const teacher = escapeHtml(lesson.преподаватель_коротко || lesson.преподаватель || '');
    const room = escapeHtml(lesson.кабинет || '');

    if (lesson.отменено) {
        return `<div class="flex items-start gap-3 p-3 rounded-2xl bg-red-500/10 border border-red-500/20 border-l-4 border-l-red-500 shadow-sm">
            <div class="w-8 h-8 rounded-lg bg-red-500 text-white flex items-center justify-center font-bold text-xs flex-shrink-0 shadow-md">${paraNum}</div>
            <div class="flex-1">
                <div class="font-bold text-sm line-through text-red-600 dark:text-red-400">${subject}</div>
                <div class="text-[10px] font-bold text-red-500 uppercase mt-0.5">Отменено</div>
            </div>
        </div>`;
    }
    if (lesson.override && lesson.override.type === 'replaced') {
        return `<div class="flex items-start gap-3 p-3 rounded-2xl bg-orange-500/10 border border-orange-500/20 border-l-4 border-l-orange-500 shadow-sm">
            <div class="w-8 h-8 rounded-lg bg-orange-500 text-white flex items-center justify-center font-bold text-xs flex-shrink-0 shadow-md">${paraNum}</div>
            <div class="flex-1">
                <div class="font-bold text-sm text-slate-900 dark:text-white">${subject}</div>
                <div class="text-[10px] text-slate-500 dark:text-slate-400 mt-1">${teacher} <br> ${room}</div>
            </div>
        </div>`;
    }
    return `<div class="flex items-start gap-3 p-3 rounded-2xl bg-white/60 dark:bg-slate-900/60 border border-white/80 dark:border-slate-800 hover:border-cyan-500/30 transition-colors shadow-sm">
        <div class="w-8 h-8 rounded-lg bg-cyan-500/10 text-cyan-600 dark:text-cyan-400 flex items-center justify-center font-bold text-xs flex-shrink-0 border border-cyan-500/20">${paraNum}</div>
        <div class="flex-1">
            <div class="font-bold text-sm text-slate-900 dark:text-white">${subject}</div>
            <div class="text-[10px] text-slate-500 dark:text-slate-400 mt-1">${teacher} <br> ${room}</div>
        </div>
    </div>`;
}