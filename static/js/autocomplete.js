(function () {
    function closePanels(exceptPanel) {
        document.querySelectorAll(".autocomplete-panel").forEach(function (panel) {
            if (panel !== exceptPanel) {
                panel.innerHTML = "";
                panel.classList.remove("is-open");
            }
        });
    }

    function renderSuggestions(input, panel, suggestions) {
        panel.innerHTML = "";

        if (!suggestions.length) {
            panel.classList.remove("is-open");
            return;
        }

        suggestions.forEach(function (suggestion) {
            var button = document.createElement("button");
            button.type = "button";
            button.className = "autocomplete-item";
            button.innerHTML =
                    "<div>" + suggestion.label + "</div>" +
                    "<small style='color:gray'>" +
                    suggestion.category +
                    "</small>";

            button.addEventListener("click", function () {
                input.value = suggestion.value;
                closePanels();
                input.form.submit();
            });

            panel.appendChild(button);
        });

        panel.classList.add("is-open");
    }

    function bindAutocomplete(input) {
        var panel = input.parentElement.querySelector(".autocomplete-panel");
        var url = input.dataset.suggestionsUrl;
        var source = input.dataset.suggestionSource || "all";
        var timeoutId = null;

        if (!panel || !url) {
            return;
        }

        input.addEventListener("input", function () {
            window.clearTimeout(timeoutId);

            timeoutId = window.setTimeout(function () {
                var query = input.value.trim();

                if (query.length < 2) {
                    renderSuggestions(input, panel, []);
                    return;
                }

                fetch(url + "?q=" + encodeURIComponent(query) + "&source=" + encodeURIComponent(source))
                    .then(function (response) {
                        return response.json();
                    })
                    .then(function (suggestions) {
                        closePanels(panel);
                        renderSuggestions(input, panel, suggestions);
                    })
                    .catch(function () {
                        renderSuggestions(input, panel, []);
                    });
            }, 160);
        });

        input.addEventListener("focus", function () {
            if (panel.children.length) {
                panel.classList.add("is-open");
            }
        });
    }

    document.addEventListener("click", function (event) {
        if (!event.target.closest(".autocomplete-wrapper")) {
            closePanels();
        }
    });

    document.addEventListener("DOMContentLoaded", function () {
        document.querySelectorAll("[data-suggestions-url]").forEach(bindAutocomplete);
    });
}());
