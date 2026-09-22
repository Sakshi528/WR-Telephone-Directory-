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
        var subcatSelect = root.querySelector("[data-org-subcategory-select]");
        if (!typeSelect) return;

        var statesUrlTemplate = root.dataset.statesUrlTemplate;
        var subcategoriesUrlTemplate = root.dataset.subcategoriesUrlTemplate;
        var organizationsUrl = root.dataset.organizationsUrl;
        var pendingState = (orgSelect && orgSelect.dataset.selectedOrgId !== undefined)
            ? (stateSelect ? stateSelect.dataset.selectedState || "" : "")
            : "";
        var pendingOrgId = orgSelect ? (orgSelect.dataset.selectedOrgId || "") : "";
        var pendingSubcategory = subcatSelect ? (subcatSelect.dataset.selectedSubcategory || "") : "";

        function isStateBased() {
            var opt = typeSelect.options[typeSelect.selectedIndex];
            return !!opt && opt.dataset.stateBased === "true";
        }

        function fetchOrganizations(typeId, state, preserveOrgId) {
            if (!orgSelect || !organizationsUrl) return;
            setPlaceholder(orgSelect, "Loading…");
            orgSelect.disabled = true;
            var url = organizationsUrl + "?type_id=" + encodeURIComponent(typeId);
            if (state) url += "&state=" + encodeURIComponent(state);
            if (subcatSelect && subcatSelect.value) url += "&subcategory=" + encodeURIComponent(subcatSelect.value);
            fetch(url)
                .then(function (r) { return r.json(); })
                .then(function (orgs) {
                    populateSelect(orgSelect, orgs, preserveOrgId);
                });
        }

        function fetchSubcategories(typeId, preserveSubcategory) {
            if (!subcatSelect || !subcategoriesUrlTemplate) return;
            setPlaceholder(subcatSelect, "Loading…");
            subcatSelect.disabled = true;
            var url = subcategoriesUrlTemplate.replace(/\/0\/subcategories/, "/" + typeId + "/subcategories");
            fetch(url)
                .then(function (r) { return r.json(); })
                .then(function (subcats) {
                    if (subcats.length) {
                        subcatSelect.parentElement.style.display = "";
                        populateSelect(subcatSelect, subcats, preserveSubcategory);
                        subcatSelect.querySelector('option[value=""]').textContent = "All Subcategories";
                    } else {
                        subcatSelect.parentElement.style.display = "none";
                        subcatSelect.value = "";
                    }
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
                    } else if (orgSelect) {
                        setPlaceholder(orgSelect, "– Select a State first –");
                        orgSelect.disabled = true;
                    }
                });
        }

        function onTypeChange(preserveState, preserveOrgId, preserveSubcategory) {
            var typeId = typeSelect.value;
            if (!typeId) {
                if (stateSelect) { stateSelect.parentElement.style.display = "none"; }
                if (subcatSelect) { subcatSelect.parentElement.style.display = "none"; subcatSelect.value = ""; }
                if (orgSelect) {
                    setPlaceholder(orgSelect, "– Select an Organization Type first –");
                    orgSelect.disabled = true;
                }
                return;
            }
            fetchSubcategories(typeId, preserveSubcategory);
            if (isStateBased() && stateSelect) {
                stateSelect.parentElement.style.display = "";
                fetchStates(typeId, preserveState, preserveOrgId);
            } else {
                if (stateSelect) { stateSelect.parentElement.style.display = "none"; }
                fetchOrganizations(typeId, null, preserveOrgId);
            }
        }

        typeSelect.addEventListener("change", function () { onTypeChange(null, null, null); });
        if (stateSelect) {
            stateSelect.addEventListener("change", function () {
                fetchOrganizations(typeSelect.value, stateSelect.value, null);
            });
        }
        if (subcatSelect) {
            subcatSelect.addEventListener("change", function () {
                fetchOrganizations(typeSelect.value, stateSelect ? stateSelect.value : null, null);
            });
        }

        if (typeSelect.value) {
            onTypeChange(pendingState, pendingOrgId, pendingSubcategory);
        } else if (stateSelect) {
            stateSelect.parentElement.style.display = "none";
        }
    }

    document.querySelectorAll("[data-org-cascade]").forEach(initCascade);
}());
