// Student Records page — searchable, filterable table + delete selected student
(function () {
    const searchInput = document.getElementById('student-search');
    const courseFilter = document.getElementById('filter-course');
    const yearFilter = document.getElementById('filter-year');
    const sectionFilter = document.getElementById('filter-section');
    const tableBody = document.getElementById('students-table-body');
    const countEl = document.getElementById('student-count');
    const totalEl = document.getElementById('total-students');
const selectAll = document.getElementById('select-all');
    const deleteBtn = document.getElementById('delete-selected-btn');
    const deleteForm = document.getElementById('delete-student-form');
    const deleteError = document.getElementById('delete-error');

    if (!searchInput || !tableBody) {
        return;
    }

    const rows = Array.prototype.slice.call(tableBody.querySelectorAll('tr'));

    // Exclude the "no students" placeholder from the real student count.
    const studentRows = rows.filter(function (row) {
        return row.id !== 'no-students-row';
    });

    function allCheckboxes() {
        return tableBody.querySelectorAll('.student-check');
    }

    function getSelectedRegisterNumbers() {
        const selected = [];
        allCheckboxes().forEach(function (cb) {
            if (cb.checked) {
                selected.push(cb.value);
            }
        });
        return selected;
    }

    function hideError() {
        if (deleteError) {
            deleteError.style.display = 'none';
        }
    }

    function showError(message) {
        if (deleteError) {
            deleteError.textContent = message;
            deleteError.style.display = 'block';
        }
    }

    function render() {
        const query = searchInput.value.trim().toLowerCase();
        const course = (courseFilter ? courseFilter.value : '').toLowerCase();
        const year = (yearFilter ? yearFilter.value : '').toLowerCase();
        const section = (sectionFilter ? sectionFilter.value : '').toLowerCase();

        let visible = 0;

        studentRows.forEach(function (row) {
            if (row.id === 'no-students-row') {
                return;
            }
            const text = row.textContent.toLowerCase();
            const matchesQuery = query === '' || text.indexOf(query) !== -1;
            const matchesCourse = course === '' || (row.dataset.course === course);
            const matchesYear = year === '' || (row.dataset.year === year);
            const matchesSection = section === '' || (row.dataset.section === section);

            const show = matchesQuery && matchesCourse && matchesYear && matchesSection;
            row.style.display = show ? '' : 'none';
            if (show) {
                visible += 1;
            }
        });

        if (countEl) {
            countEl.textContent = visible + ' of ' + studentRows.length + ' student(s)';
        }
        if (totalEl) {
            totalEl.textContent = studentRows.length;
        }

        // Keep "select all" in sync with the visible (filtered) checkboxes.
        if (selectAll) {
            const visibleBoxes = Array.prototype.filter.call(allCheckboxes(), function (cb) {
                return cb.closest('tr').style.display !== 'none';
            });
            if (visibleBoxes.length > 0) {
                selectAll.checked = visibleBoxes.every(function (cb) { return cb.checked; });
                selectAll.indeterminate = !selectAll.checked && visibleBoxes.some(function (cb) { return cb.checked; });
            } else {
                selectAll.checked = false;
                selectAll.indeterminate = false;
            }
        }
    }

    // Select-all header checkbox toggles all visible (filtered) row checkboxes.
    if (selectAll) {
        selectAll.addEventListener('change', function () {
            const shouldCheck = selectAll.checked;
            allCheckboxes().forEach(function (cb) {
                if (cb.closest('tr').style.display !== 'none') {
                    cb.checked = shouldCheck;
                }
            });
        });
    }

    // Any checkbox change re-evaluates the select-all state.
    tableBody.addEventListener('change', function (event) {
        if (event.target && event.target.classList && event.target.classList.contains('student-check')) {
            render();
        }
    });

// Delete Selected Student button.
    if (deleteBtn && deleteForm) {
        deleteBtn.addEventListener('click', function () {
            hideError();
            const selected = getSelectedRegisterNumbers();
            if (selected.length === 0) {
                showError('Please select at least one student to delete.');
                return;
            }
            const confirmed = window.confirm(
                'Are you sure you want to delete ' + selected.length + ' selected student(s)? This cannot be undone.'
            );
            if (!confirmed) {
                return;
            }
            // Remove any previously-added hidden register inputs.
            const existing = deleteForm.querySelectorAll('input[name="register_number"]');
            existing.forEach(function (el) {
                el.remove();
            });
            // Add one hidden input per selected student.
            selected.forEach(function (registerNumber) {
                const input = document.createElement('input');
                input.type = 'hidden';
                input.name = 'register_number';
                input.value = registerNumber;
                deleteForm.appendChild(input);
            });
            deleteForm.submit();
        });
    }

    searchInput.addEventListener('input', render);
    if (courseFilter) { courseFilter.addEventListener('change', render); }
    if (yearFilter) { yearFilter.addEventListener('change', render); }
    if (sectionFilter) { sectionFilter.addEventListener('change', render); }

    render();
})();
