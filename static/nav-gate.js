// nav-gate.js — sembunyikan menu disposisi untuk yang tak berhak memantau.
// Hak efektif dari backend (/api/user/profile -> disp_pantau, diatur admin).
// Keamanan tetap di backend (403); ini hanya agar opsinya tidak terlihat.
document.addEventListener('DOMContentLoaded', () => {
  fetch('/api/user/profile').then(r => r.json()).then(u => {
    const boleh = ('disp_pantau' in u) ? !!u.disp_pantau
      : (u.role === 'admin' || u.role === 'atasan' || u.bagian === 'persuratan');
    if (!boleh) {
      document.querySelectorAll('a[href="riwayat-disposisi.html"]').forEach(el => {
        el.style.display = 'none';
      });
    }
  }).catch(() => {});
});
