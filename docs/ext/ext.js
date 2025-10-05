document.addEventListener('DOMContentLoaded', () => {
    const as = document.querySelectorAll('a[href*=".md#"]');
    as.forEach(a => {
        try {
            const url = new URL(a.href, window.location.origin);
            if (url.pathname.endsWith('.md')) {
                url.pathname = url.pathname.replace(/\.md$/, '.html');
                a.href = url.toString();
            }
        } catch (e) { /* ignore */ }
    });
});
