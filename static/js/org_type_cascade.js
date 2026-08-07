/*
 * Reusable Organization Type -> State -> Organization cascading dropdown.
 *
 * Usage: wrap the three <select> elements in a container carrying
 * data-org-cascade, data-states-url-template ("/api/organization-types/0/states",
 * the "0" is replaced with the chosen type id) and data-organizations-url
 * ("/api/organizations"). Mark each select with data-org-type-select /
 * data-org-state-select / data-org-select. Each <option> on the type select
 * needs data-state-based="true"/"false" (whether picking that type reveals
 * the State dropdown). For edit forms / filter forms that should restore a
 * prior selection on load, set data-selected-state on the state select and
 * data-selected-org-id on the organization select.
 */
(function () {
    function setPlaceholder(select, text) {
        select.innerHTML = "";
        var opt = document.createElement("option");
        opt.value = "";
        opt.textContent = text;
        select.appendChild(opt);
    }

    function populateSelect(select, items, selectedValue) {
        select.innerHTML = "";
        var placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "– Select –";
        select.appendChild(placeholder);
        items.forEach(function (item) {
            var opt = document.createElement("option");
            opt.value = item.id;
            opt.textContent = item.name;
            if (selectedValue !== undefined && selectedValue !== null &&
                    String(item.id) === String(selectedValue)) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });
        select.disabled = false;
    }

    function initCascade(root) {
        var typeSelect = root.querySelector("[data-org-type-select]");
        var stateSelect = root.querySelector("[data-org-state-select]");
        var orgSelect = root.querySelector("[data-org-select]");
        if (!typeSelect || !orgSelect) return;

        var statesUrlTemplate = root.dataset.statesUrlTemplate;
        var organizationsUrl = root.dataset.organizationsUrl;
        var pendingState = orgSelect.dataset.selectedOrgId !== undefined
            ? (stateSelect ? stateSelect.dataset.selectedState || "" : "")
            : "";
        var pendingOrgId = orgSelect.dataset.selectedOrgId || "";

        function isStateBased() {
            var opt = typeSelect.options[typeSelect.selectedIndex];
            return !!opt && opt.dataset.stateBased === "true";
        }

        function fetchOrganizations(typeId, state, preserveOrgId) {
            setPlaceholder(orgSelect, "Loading…");
            orgSelect.disabled = true;
            var url = organizationsUrl + "?type_id=" + encodeURIComponent(typeId);
            if (state) url += "&state=" + encodeURIComponent(state);
            fetch(url)
                .then(function (r) { return r.json(); })
                .then(function (orgs) {
                    populateSelect(orgSelect, orgs, preserveOrgId);
                });
        }

        function fetchStates(typeId, preserveState, preserveOrgId) {
            setPlaceholder(stateSelect, "Loading…");
            stateSelect.disabled = true;
            var url = statesUrlTemplate.replace(/\/0\/states/, "/" + typeId + "/states");
            fetch(url)
                .then(function (r) { return r.json(); })
                .then(function (states) {
                    populateSelect(stateSelect, states, preserveState);
                    if (preserveState) {
                        fetchOrganizations(typeId, preserveState, preserveOrgId);
                    } else {
                        setPlaceholder(orgSelect, "– Select a State first –");
                        orgSelect.disabled = true;
                    }
                });
        }

        function onTypeChange(preserveState, preserveOrgId) {
            var typeId = typeSelect.value;
            if (!typeId) {
                if (stateSelect) { stateSelect.parentElement.style.display = "none"; }
                setPlaceholder(orgSelect, "– Select an Organization Type first –");
                orgSelect.disabled = true;
                return;
            }
            if (isStateBased() && stateSelect) {
                stateSelect.parentElement.style.display = "";
                fetchStates(typeId, preserveState, preserveOrgId);
            } else {
                if (stateSelect) { stateSelect.parentElement.style.display = "none"; }
                fetchOrganizations(typeId, null, preserveOrgId);
            }
        }

        typeSelect.addEventListener("change", function () { onTypeChange(null, null); });
        if (stateSelect) {
            stateSelect.addEventListener("change", function () {
                fetchOrganizations(typeSelect.value, stateSelect.value, null);
            });
        }

        if (typeSelect.value) {
            onTypeChange(pendingState, pendingOrgId);
        } else if (stateSelect) {
            stateSelect.parentElement.style.display = "none";
        }
    }

    document.querySelectorAll("[data-org-cascade]").forEach(initCascade);
}());
