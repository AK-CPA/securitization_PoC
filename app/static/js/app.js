/* ODRT Application JavaScript */

// Add loading spinner to forms on submit
document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('form').forEach(function(form) {
        form.addEventListener('submit', function() {
            const btn = form.querySelector('button[type="submit"]');
            if (btn && !btn.disabled) {
                btn.disabled = true;
                const originalText = btn.innerHTML;
                btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2" role="status"></span>Processing...';
                // Re-enable after 30 seconds as a safety net
                setTimeout(function() {
                    btn.disabled = false;
                    btn.innerHTML = originalText;
                }, 30000);
            }
        });
    });
});
