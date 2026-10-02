const loginMenu = document.querySelector('.login-menu');
const loginButton = document.querySelector('.login-button');

if (loginMenu && loginButton) {
    loginButton.addEventListener('click', () => {
        const isOpen = loginMenu.classList.toggle('open');
        loginButton.setAttribute('aria-expanded', String(isOpen));
    });

    document.addEventListener('click', (event) => {
        if (!loginMenu.contains(event.target)) {
            loginMenu.classList.remove('open');
            loginButton.setAttribute('aria-expanded', 'false');
        }
    });
}

const passwordInput = document.querySelector('#password');
const passwordToggle = document.querySelector('.password-toggle');

if (passwordInput && passwordToggle) {
    passwordToggle.addEventListener('click', () => {
        const isPassword = passwordInput.type === 'password';
        passwordInput.type = isPassword ? 'text' : 'password';
        passwordToggle.setAttribute('aria-label', isPassword ? 'Hide password' : 'Show password');
        passwordToggle.classList.toggle('visible', isPassword);
    });
}

const sidebar = document.querySelector('.dashboard-sidebar');
const sidebarToggle = document.querySelector('.sidebar-toggle');

if (sidebar && sidebarToggle) {
    sidebarToggle.addEventListener('click', () => {
        sidebar.classList.toggle('open');
    });
}

const studentSort = document.querySelector('#sort');
const studentTableBody = document.querySelector('.student-table tbody');

if (studentSort && studentTableBody) {
    const sortStudents = () => {
        const rows = Array.from(studentTableBody.querySelectorAll('tr'));
        if (rows.length < 2) {
            return;
        }

        rows.sort((firstRow, secondRow) => {
            const firstName = firstRow.cells[0].textContent.trim().toLocaleLowerCase();
            const secondName = secondRow.cells[0].textContent.trim().toLocaleLowerCase();
            const comparison = firstName.localeCompare(secondName);
            return studentSort.value === 'desc' ? -comparison : comparison;
        });

        rows.forEach((row) => studentTableBody.appendChild(row));
        const url = new URL(window.location.href);
        url.searchParams.set('sort', studentSort.value);
        window.history.replaceState({}, '', url);
    };

    studentSort.addEventListener('change', sortStudents);
}

// Render SGPA line chart when data is available
document.addEventListener('DOMContentLoaded', () => {
    try {
        const sgpa = window.SGPA_DATA;
        const ctx = document.getElementById('sgpaChart');
        if (ctx && Array.isArray(sgpa) && sgpa.length) {
            const labels = sgpa.map((_, i) => `Sem ${i+1}`);
            const gradient = ctx.getContext('2d').createLinearGradient(0, 0, 0, 160);
            gradient.addColorStop(0, 'rgba(121,183,152,0.28)');
            gradient.addColorStop(1, 'rgba(44,111,71,0.02)');

            new Chart(ctx.getContext('2d'), {
                type: 'line',
                data: {
                    labels,
                    datasets: [{
                        label: 'SGPA',
                        data: sgpa,
                        fill: true,
                        backgroundColor: gradient,
                        borderColor: '#2c6f47',
                        pointBackgroundColor: '#ffffff',
                        pointBorderColor: '#2c6f47',
                        pointRadius: 5,
                        tension: 0.35,
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        y: {
                            beginAtZero: false,
                            suggestedMin: Math.min(...sgpa) - 0.5,
                            suggestedMax: Math.max(...sgpa) + 0.5,
                            grid: { color: 'rgba(34,64,48,0.06)' },
                            ticks: { color: '#557064' }
                        },
                        x: {
                            grid: { display: false },
                            ticks: { color: '#6a7a69' }
                        }
                    },
                    plugins: {
                        legend: { display: false },
                        tooltip: { mode: 'index', intersect: false }
                    }
                }
            });
        }
    } catch (e) {
        // ignore chart errors
        console.error('SGPA chart render failed', e);
    }
});

// Attendance changes save immediately; no separate Save button is required.
document.querySelectorAll('[data-attendance-form]').forEach((container) => {
    const getContext = () => {
        const form = container.matches('form') ? container : null;
        return {
            course: container.dataset.course || form?.elements.course?.value,
            year: container.dataset.year || form?.elements.year?.value,
            semester: container.dataset.semester || form?.elements.semester?.value,
            subject: container.dataset.subject || form?.elements.subject?.value,
            attendance_date: container.dataset.attendanceDate || form?.elements.attendance_date?.value,
        };
    };
    container.querySelectorAll('input[type="radio"][name^="status_"]').forEach((radio) => {
        radio.addEventListener('change', async () => {
            const context = getContext();
            const notice = document.querySelector('.attendance-autosave');
            const register_number = radio.name.replace('status_', '');
            radio.closest('.status-options')?.querySelectorAll('.status-option').forEach((label) => label.classList.remove('selected'));
            radio.closest('label')?.classList.add('selected');
            if (notice) notice.textContent = 'Saving attendance…';
            try {
                const response = await fetch('/attendance/save', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({...context, register_number, status: radio.value}),
                });
                const result = await response.json();
                if (!response.ok || !result.ok) throw new Error(result.message || 'Save failed');
                if (notice) notice.textContent = '✓ Saved automatically and synchronized to Excel.';
            } catch (error) {
                if (notice) notice.textContent = '✕ ' + error.message + ' Please try again.';
            }
        });
    });
    if (container.matches('form')) {
        container.querySelectorAll('select, input[type="date"]').forEach((field) => {
            field.addEventListener('change', () => container.submit());
        });
    }
});
