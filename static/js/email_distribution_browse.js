(function () {
    var app = document.getElementById("edbApp");
    if (!app) return;

    var dataUrl = app.dataset.dataUrl;
    var exportXlsxUrl = app.dataset.exportXlsxUrl;
    var exportCsvUrl = app.dataset.exportCsvUrl;
    var mailtoLimit = parseInt(app.dataset.mailtoLimit, 10);
    var statesUrlTemplate = app.dataset.statesUrlTemplate;

    var isAdmin = app.dataset.isAdmin === "true";
    var editUrlTemplates = {
        employee: app.dataset.editUrlEmployee,
        directory_number: app.dataset.editUrlDirectorynumber,
        administrative_head: app.dataset.editUrlAdministrativehead,
    };
    var deleteUrlTemplates = {
        employee: app.dataset.deleteUrlEmployee,
        directory_number: app.dataset.deleteUrlDirectorynumber,
    };
    function actionUrl(templates, recordType, id) {
        var tmpl = templates[recordType];
        return tmpl ? tmpl.replace(/\/0(?:\?.*)?$/, "/" + id) : null;
    }

    // Category is fixed for this page load (set by the card that linked
    // here) rather than user-selectable -- see routes.user_routes.email_distribution_browse.
    var category = app.dataset.category;
    var categoryId = app.dataset.categoryId;
    var categoryIsStateBased = app.dataset.stateBased === "true";

    var stateWrapEl = document.getElementById("edbStateWrap");
    var stateSelect = document.getElementById("edbState");
    var subcategorySelect = document.getElementById("edbSubcategory");
    var contactTypeSelect = document.getElementById("edbContactType");
    var searchInput = document.getElementById("edbSearch");

    var placeholderEl = document.getElementById("edbPlaceholder");
    var gridWrapEl = document.getElementById("edbGridWrap");
    var gridBodyEl = document.getElementById("edbGridBody");
    var noResultsEl = document.getElementById("edbNoResults");
    var actionsEl = document.getElementById("edbActions");
    var selectAllEl = document.getElementById("edbSelectAll");
    var noResultsTextEl = document.getElementById("edbNoResultsText");
    var showAllBtn = document.getElementById("edbShowAllBtn");

    var fallbackModalEl = document.getElementById("edbMailtoFallbackModal");
    var fallbackCountEl = document.getElementById("edbFallbackCount");

    var allRows = [];      // full result set for the current category + contact type
    var visibleRows = [];  // allRows after the live search filter

    function escapeHtml(s) {
        var div = document.createElement("div");
        div.textContent = s || "";
        return div.innerHTML;
    }

    function currentQuery() {
        return {
            category: category,
            state: stateWrapEl.style.display !== "none" ? stateSelect.value : "",
            subcategory: subcategorySelect ? subcategorySelect.value : "",
            contact_type: contactTypeSelect.value,
            q: searchInput.value.trim(),
        };
    }

    function buildUrl(base, params) {
        var usp = new URLSearchParams();
        Object.keys(params).forEach(function (k) {
            if (params[k]) usp.set(k, params[k]);
        });
        return base + "?" + usp.toString();
    }

    function updateStateVisibility() {
        if (categoryIsStateBased) {
            stateWrapEl.style.display = "";
            stateSelect.innerHTML = '<option value="">Loading…</option>';
            fetch(statesUrlTemplate.replace(/\/0\/states/, "/" + categoryId + "/states"))
                .then(function (r) { return r.json(); })
                .then(function (states) {
                    stateSelect.innerHTML = '<option value="">All States</option>';
                    states.forEach(function (s) {
                        var o = document.createElement("option");
                        o.value = s.id;
                        o.textContent = s.name;
                        stateSelect.appendChild(o);
                    });
                });
        } else {
            stateWrapEl.style.display = "none";
            stateSelect.innerHTML = '<option value="">All States</option>';
        }
    }

    function fetchAndRender() {
        var url = buildUrl(dataUrl, {
            category: category, contact_type: contactTypeSelect.value,
            state: categoryIsStateBased ? stateSelect.value : "",
            subcategory: subcategorySelect ? subcategorySelect.value : "",
        });
        fetch(url)
            .then(function (r) { return r.json(); })
            .then(function (data) {
                placeholderEl.style.display = "none";
                allRows = data.rows || [];
                if (data.summary) {
                    document.getElementById("edbSumOrgs").textContent = data.summary.organizations;
                    document.getElementById("edbSumEmployees").textContent = data.summary.employees;
                }
                applySearch();
            });
    }

    function applySearch() {
        var q = searchInput.value.trim().toLowerCase();
        visibleRows = !q ? allRows.slice() : allRows.filter(function (r) {
            return (r.organization || "").toLowerCase().indexOf(q) !== -1 ||
                   (r.name || "").toLowerCase().indexOf(q) !== -1 ||
                   (r.designation || "").toLowerCase().indexOf(q) !== -1 ||
                   (r.email || "").toLowerCase().indexOf(q) !== -1;
        });
        renderGrid();
    }

    function renderGrid() {
        gridBodyEl.innerHTML = "";
        selectAllEl.checked = false;

        if (!visibleRows.length) {
            gridWrapEl.style.display = "none";
            noResultsEl.style.display = "";
            actionsEl.style.display = allRows.length ? "" : "none";
            // A narrowed Contact Type with zero results usually just means this
            // category has no one in that role (e.g. no Administrative Heads at
            // a private RE Generator) -- offer a one-click way back to everyone
            // instead of a bare "no results".
            if (contactTypeSelect.value !== "all" && !searchInput.value.trim()) {
                noResultsTextEl.textContent = "No " +
                    contactTypeSelect.options[contactTypeSelect.selectedIndex].text +
                    " found for this category.";
                showAllBtn.style.display = "";
            } else {
                noResultsTextEl.textContent = "No contacts match the current filters.";
                showAllBtn.style.display = "none";
            }
            return;
        }

        noResultsEl.style.display = "none";
        gridWrapEl.style.display = "";
        actionsEl.style.display = "";

        visibleRows.forEach(function (row, idx) {
            var tr = document.createElement("tr");
            tr.innerHTML =
                '<td class="text-center"><input type="checkbox" class="edb-row-check" data-idx="' + idx + '"></td>' +
                "<td>" + escapeHtml(row.organization) + "</td>" +
                "<td>" + escapeHtml(row.name) + "</td>" +
                "<td>" + escapeHtml(row.designation) + "</td>" +
                "<td>" + (row.email ? '<a href="mailto:' + escapeHtml(row.email) + '">' + escapeHtml(row.email) + "</a>" : "–") + "</td>" +
                (isAdmin ? "<td class='text-center text-nowrap'>" + rowActionsHtml(row) + "</td>" : "");
            gridBodyEl.appendChild(tr);
        });
    }

    function rowActionsHtml(row) {
        if (!row.record_id || !row.record_type) return "–";
        var editUrl = actionUrl(editUrlTemplates, row.record_type, row.record_id);
        var deleteUrl = actionUrl(deleteUrlTemplates, row.record_type, row.record_id);
        var html = "";
        if (editUrl) {
            html += '<a href="' + editUrl + '" class="btn btn-outline-warning btn-sm" title="Edit">' +
                '<i class="bi bi-pencil"></i></a> ';
        }
        if (deleteUrl) {
            html += '<button type="button" class="btn btn-outline-danger btn-sm edb-delete-btn" ' +
                'data-delete-url="' + deleteUrl + '" data-name="' + escapeHtml(row.name) + '" title="Delete">' +
                '<i class="bi bi-trash"></i></button>';
        }
        return html || "–";
    }

    function selectedRows() {
        var checks = gridBodyEl.querySelectorAll(".edb-row-check:checked");
        return Array.prototype.map.call(checks, function (c) { return visibleRows[parseInt(c.dataset.idx, 10)]; });
    }

    function dedupeSortedEmails(rows) {
        var seen = {};
        rows.forEach(function (r) {
            var email = (r.email || "").trim();
            if (!email) return;
            var key = email.toLowerCase();
            if (!(key in seen)) seen[key] = email;
        });
        return Object.keys(seen).map(function (k) { return seen[k]; }).sort(function (a, b) {
            return a.toLowerCase().localeCompare(b.toLowerCase());
        });
    }

    function copyText(text, button) {
        navigator.clipboard.writeText(text).then(function () {
            if (!button) return;
            var original = button.innerHTML;
            button.innerHTML = '<i class="bi bi-check2 me-1"></i>Copied!';
            setTimeout(function () { button.innerHTML = original; }, 1500);
        });
    }

    document.getElementById("edbCopySelected").addEventListener("click", function (e) {
        copyText(dedupeSortedEmails(selectedRows()).join(";\n"), e.currentTarget);
    });

    document.getElementById("edbCopyAll").addEventListener("click", function (e) {
        copyText(dedupeSortedEmails(visibleRows).join(";\n"), e.currentTarget);
    });

    document.getElementById("edbCopyNameEmail").addEventListener("click", function (e) {
        var rows = selectedRows().length ? selectedRows() : visibleRows;
        var lines = rows.filter(function (r) { return r.email; })
            .map(function (r) { return (r.name || "") + " <" + r.email + ">"; });
        copyText(lines.join("\n"), e.currentTarget);
    });

    document.getElementById("edbCopyFullDetails").addEventListener("click", function (e) {
        var rows = selectedRows().length ? selectedRows() : visibleRows;
        var blocks = rows.map(function (r) {
            return "Organization: " + r.organization +
                "\nName: " + r.name +
                "\nRole / Designation: " + r.designation +
                "\nEmail: " + (r.email || "-") +
                "\nMobile: " + (r.mobile || "-");
        });
        copyText(blocks.join("\n\n"), e.currentTarget);
    });

    document.getElementById("edbCompose").addEventListener("click", function () {
        var emails = dedupeSortedEmails(visibleRows);
        var mailto = "mailto:" + encodeURIComponent(emails.join(","));
        if (mailto.length <= mailtoLimit) {
            window.location.href = mailto;
        } else {
            fallbackCountEl.textContent = emails.length;
            bootstrap.Modal.getOrCreateInstance(fallbackModalEl).show();
        }
    });

    document.getElementById("edbFallbackCopyBtn").addEventListener("click", function (e) {
        copyText(dedupeSortedEmails(visibleRows).join(";\n"), e.currentTarget);
    });

    selectAllEl.addEventListener("change", function () {
        var checks = gridBodyEl.querySelectorAll(".edb-row-check");
        checks.forEach(function (c) { c.checked = selectAllEl.checked; });
    });

    function updateExportLinks() {
        var q = currentQuery();
        document.getElementById("edbExportXlsx").href = buildUrl(exportXlsxUrl, q);
        document.getElementById("edbExportCsv").href = buildUrl(exportCsvUrl, q);
    }

    stateSelect.addEventListener("change", function () { fetchAndRender(); updateExportLinks(); });
    if (subcategorySelect) {
        subcategorySelect.addEventListener("change", function () { fetchAndRender(); updateExportLinks(); });
    }
    contactTypeSelect.addEventListener("change", function () { fetchAndRender(); updateExportLinks(); });
    showAllBtn.addEventListener("click", function () {
        contactTypeSelect.value = "all";
        fetchAndRender();
        updateExportLinks();
    });
    searchInput.addEventListener("input", function () { applySearch(); updateExportLinks(); });

    // Delete buttons are created dynamically in renderGrid() -- delegate from
    // the (static) table body rather than binding one listener per row.
    gridBodyEl.addEventListener("click", function (e) {
        var btn = e.target.closest(".edb-delete-btn");
        if (!btn) return;
        if (!confirm("Delete \"" + btn.dataset.name + "\"? This cannot be undone.")) return;
        var form = document.createElement("form");
        form.method = "POST";
        form.action = btn.dataset.deleteUrl;
        document.body.appendChild(form);
        form.submit();
    });

    updateStateVisibility();
    fetchAndRender();
    updateExportLinks();
}());
