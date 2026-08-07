(function () {
    var container = document.getElementById("edGroupActions");
    if (!container) return;

    var emailsUrl = container.dataset.emailsUrl;
    var mailtoLimit = parseInt(container.dataset.mailtoLimit, 10);
    var rows = JSON.parse(container.dataset.rows || "[]");

    var copyBtn = document.getElementById("edCopyBtn");
    var composeBtn = document.getElementById("edComposeBtn");
    var copySelectedBtn = document.getElementById("edCopySelected");
    var copyNameEmailBtn = document.getElementById("edCopyNameEmail");
    var copyFullDetailsBtn = document.getElementById("edCopyFullDetails");
    var selectAllEl = document.getElementById("edSelectAll");
    var gridBodyEl = document.getElementById("edGridBody");
    var fallbackModalEl = document.getElementById("edMailtoFallbackModal");
    var fallbackCount = document.getElementById("edFallbackCount");
    var fallbackCopyBtn = document.getElementById("edFallbackCopyBtn");

    var cachedData = null;

    function fetchEmails() {
        if (cachedData) return Promise.resolve(cachedData);
        return fetch(emailsUrl)
            .then(function (r) { return r.json(); })
            .then(function (data) {
                cachedData = data;
                return data;
            });
    }

    function copyToClipboard(text, button) {
        navigator.clipboard.writeText(text).then(function () {
            var original = button.innerHTML;
            button.innerHTML = '<i class="bi bi-check2 me-1"></i>Copied!';
            setTimeout(function () { button.innerHTML = original; }, 1500);
        });
    }

    function selectedRows() {
        if (!gridBodyEl) return [];
        var checks = gridBodyEl.querySelectorAll(".ed-row-check:checked");
        return Array.prototype.map.call(checks, function (c) { return rows[parseInt(c.dataset.idx, 10)]; });
    }

    function dedupeSortedEmails(rowList) {
        var seen = {};
        rowList.forEach(function (r) {
            var email = (r.email || "").trim();
            if (!email) return;
            var key = email.toLowerCase();
            if (!(key in seen)) seen[key] = email;
        });
        return Object.keys(seen).map(function (k) { return seen[k]; }).sort(function (a, b) {
            return a.toLowerCase().localeCompare(b.toLowerCase());
        });
    }

    copyBtn.addEventListener("click", function () {
        fetchEmails().then(function (data) { copyToClipboard(data.emails.join(";\n"), copyBtn); });
    });

    fallbackCopyBtn.addEventListener("click", function () {
        fetchEmails().then(function (data) { copyToClipboard(data.emails.join(";\n"), fallbackCopyBtn); });
    });

    composeBtn.addEventListener("click", function () {
        fetchEmails().then(function (data) {
            if (data.mailto_safe) {
                window.location.href = data.mailto;
            } else {
                fallbackCount.textContent = data.email_count;
                var modal = bootstrap.Modal.getOrCreateInstance(fallbackModalEl);
                modal.show();
            }
        });
    });

    if (copySelectedBtn) {
        copySelectedBtn.addEventListener("click", function () {
            copyToClipboard(dedupeSortedEmails(selectedRows()).join(";\n"), copySelectedBtn);
        });
    }

    if (copyNameEmailBtn) {
        copyNameEmailBtn.addEventListener("click", function () {
            var rowList = selectedRows().length ? selectedRows() : rows;
            var lines = rowList.filter(function (r) { return r.email; })
                .map(function (r) { return (r.name || "") + " <" + r.email + ">"; });
            copyToClipboard(lines.join("\n"), copyNameEmailBtn);
        });
    }

    if (copyFullDetailsBtn) {
        copyFullDetailsBtn.addEventListener("click", function () {
            var rowList = selectedRows().length ? selectedRows() : rows;
            var blocks = rowList.map(function (r) {
                return "Organization: " + r.organization +
                    "\nName: " + r.name +
                    "\nDesignation: " + r.designation +
                    "\nEmail: " + (r.email || "-") +
                    "\nMobile: " + (r.mobile || "-");
            });
            copyToClipboard(blocks.join("\n\n"), copyFullDetailsBtn);
        });
    }

    if (selectAllEl && gridBodyEl) {
        selectAllEl.addEventListener("change", function () {
            var checks = gridBodyEl.querySelectorAll(".ed-row-check");
            checks.forEach(function (c) { c.checked = selectAllEl.checked; });
        });
    }
}());
