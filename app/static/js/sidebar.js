
/* =====================================================
   OM ENTERPRISE ERP — SIDEBAR CONTROLS
   File: app/static/js/sidebar.js
   ===================================================== */

document.addEventListener("DOMContentLoaded", function () {
    const appShell = document.querySelector(".app-shell");
    const sidebar = document.getElementById("appSidebar");
    const toggleButton = document.getElementById("sidebarToggle");

    if (!appShell || !sidebar || !toggleButton) {
        return;
    }

    // Mobile overlay: click outside sidebar to close it.
    let overlay = document.querySelector(".sidebar-overlay");

    if (!overlay) {
        overlay = document.createElement("div");
        overlay.className = "sidebar-overlay";
        overlay.setAttribute("aria-hidden", "true");
        appShell.appendChild(overlay);
    }

    function isMobile() {
        return window.matchMedia("(max-width: 760px)").matches;
    }

    function openMobileSidebar() {
        appShell.classList.add("mobile-sidebar-open");
        toggleButton.setAttribute("aria-expanded", "true");
        sidebar.setAttribute("aria-hidden", "false");
    }

    function closeMobileSidebar() {
        appShell.classList.remove("mobile-sidebar-open");
        toggleButton.setAttribute("aria-expanded", "false");
        sidebar.setAttribute("aria-hidden", "true");
    }

    function toggleSidebar() {
        if (isMobile()) {
            if (appShell.classList.contains("mobile-sidebar-open")) {
                closeMobileSidebar();
            } else {
                openMobileSidebar();
            }
            return;
        }

        appShell.classList.toggle("sidebar-collapsed");

        const isCollapsed =
            appShell.classList.contains("sidebar-collapsed");

        toggleButton.setAttribute(
            "aria-expanded",
            String(!isCollapsed)
        );
    }

    toggleButton.addEventListener("click", toggleSidebar);

    overlay.addEventListener("click", closeMobileSidebar);

    // Close mobile sidebar after choosing a real navigation link.
    sidebar.querySelectorAll(".nav-link").forEach(function (link) {
        link.addEventListener("click", function () {
            if (isMobile() && link.getAttribute("href") !== "#") {
                closeMobileSidebar();
            }
        });
    });

    // Escape key closes the mobile sidebar.
    document.addEventListener("keydown", function (event) {
        if (
            event.key === "Escape" &&
            appShell.classList.contains("mobile-sidebar-open")
        ) {
            closeMobileSidebar();
            toggleButton.focus();
        }
    });

    // Reset mobile drawer state when switching to desktop.
    window.addEventListener("resize", function () {
        if (!isMobile()) {
            closeMobileSidebar();
            sidebar.removeAttribute("aria-hidden");
        }
    });

    // Set initial accessibility state.
    if (isMobile()) {
        sidebar.setAttribute("aria-hidden", "true");
        toggleButton.setAttribute("aria-expanded", "false");
    } else {
        sidebar.removeAttribute("aria-hidden");
        toggleButton.setAttribute("aria-expanded", "true");
    }
});
