"""Real-browser checks for the shared visual layer with all SFC requests mocked."""

from contextlib import contextmanager
import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_dashboard_server as server
import sfc_report_tools as reports
from werkzeug.serving import make_server


@contextmanager
def isolated_server():
    tools = reports.registry(include_inputs=False)
    with (
        tempfile.TemporaryDirectory() as folder,
        patch.object(reports, "RUN_ROOT", Path(folder)),
        patch.object(reports, "registry", return_value=tools),
        patch.object(server.sfc_command_jobs, "latest_prompt", return_value=""),
        patch.object(server.sfc_command_jobs, "latest_job_id", return_value=""),
        patch.object(server.sfc_command_jobs, "saved_results", return_value=[]),
    ):
        http = make_server("127.0.0.1", 0, server.app, threaded=True)
        thread = threading.Thread(target=http.serve_forever, daemon=True)
        thread.start()
        try:
            yield f"http://127.0.0.1:{http.server_port}"
        finally:
            http.shutdown()
            thread.join(timeout=5)


def capture(page, name, full_page=True):
    folder = os.environ.get("CHINGLUH_DESIGN_SCREENSHOTS")
    if folder:
        path = Path(folder)
        path.mkdir(parents=True, exist_ok=True)
        page.screenshot(
            path=str(path / name), full_page=full_page, animations="disabled"
        )


@unittest.skipUnless(
    os.environ.get("CHINGLUH_UI_TESTS") == "1",
    "Set CHINGLUH_UI_TESTS=1 for headless Chrome",
)
class WorkspaceAppearanceTests(unittest.TestCase):
    def test_command_panel_is_clear_and_collapsed_on_reload(self):
        from playwright.sync_api import sync_playwright, expect

        with (
            isolated_server() as url,
            patch.object(
                server.sfc_command_jobs,
                "latest_prompt",
                return_value="Previous MQAA command",
            ),
            patch.object(
                server.sfc_command_jobs, "latest_job_id", return_value="previous-job"
            ),
            patch.object(
                server.sfc_command_jobs,
                "saved_results",
                return_value=[
                    dict(
                        files=[
                            dict(
                                name="Saved report.xlsx",
                                url="/api/exports/fixture/report.xlsx",
                            )
                        ]
                    )
                ],
            ),
            sync_playwright() as playwright,
        ):
            browser = playwright.chromium.launch(channel="chrome", headless=True)
            context = browser.new_context(viewport=dict(width=1440, height=1000))
            context.add_init_script(
                "localStorage.setItem('sfc-last-prompt','Old browser command');localStorage.setItem('sfc-last-job','previous-job');"
            )
            page = context.new_page()
            posts = []

            def api(route):
                request = route.request
                if request.method == "POST":
                    posts.append(request.url)
                    route.fulfill(
                        content_type="text/event-stream",
                        body='data: {"step":5,"status":"done","msg":"Completed fixture","job_id":"new-job"}\n\n',
                    )
                    return
                path = request.url.split("/api/", 1)[-1]
                data = {"success": True}
                if path.startswith("command/jobs/"):
                    data = {
                        "id": "previous-job",
                        "status": "done",
                        "prompt": "Previous MQAA command",
                    }
                elif path.startswith("status"):
                    data = {"connected": False}
                elif path == "page/inspect":
                    data = {"success": False, "error": "Offline fixture"}
                elif path == "factories":
                    data = {"factories": []}
                elif path == "memory":
                    data = {"modules": [], "qms_modules": []}
                elif path == "nav/config":
                    data = {"success": True, "tiles": server.DEFAULT_NAV_TILES}
                route.fulfill(json=data)

            page.route("**/api/**", api)
            page.goto(url)
            expect(page.locator("#ai-prompt-input")).to_have_value("")
            expect(page.locator("#ai-exec-console")).not_to_be_visible()
            expect(page.locator("#resume-command-progress")).to_be_hidden()
            self.assertFalse(
                page.locator("#saved-command-results").evaluate("node=>node.open")
            )
            self.assertIsNone(page.evaluate("localStorage.getItem('sfc-last-prompt')"))
            self.assertLess(
                page.locator(".ai-prompt-box").bounding_box()["height"], 280
            )
            self.assertEqual(posts, [])
            folder = os.environ.get("CHINGLUH_DESIGN_SCREENSHOTS")
            if folder:
                page.locator(".ai-prompt-box").screenshot(
                    path=str(Path(folder) / "command-panel-compact.png"),
                    animations="disabled",
                )
            page.locator("#saved-command-results summary").click()
            expect(page.get_by_role("link", name="↓ Saved report.xlsx")).to_be_visible()
            page.locator(".quick-chips").get_by_text(
                "QA HFPA (Q504)", exact=True
            ).click()
            expect(page.locator("#ai-prompt-input")).to_have_value("Vào HFPA")
            expect(page.locator("#exec-status-badge")).to_contain_text("HOÀN THÀNH TOÀN BỘ")
            expect(page.locator("#ai-exec-console")).to_be_visible()
            page.get_by_role("button", name="Thu gọn tiến trình", exact=True).click()
            expect(page.locator("#ai-exec-console")).not_to_be_visible()
            page.locator("#resume-command-progress").click()
            expect(page.locator("#ai-exec-console")).to_be_visible()
            self.assertEqual(
                len(posts), 1, f"Reopening progress must not resubmit the command: {posts}"
            )
            page.reload()
            expect(page.locator("#ai-prompt-input")).to_have_value("")
            expect(page.locator("#ai-exec-console")).not_to_be_visible()
            self.assertFalse(
                page.locator("#saved-command-results").evaluate("node=>node.open")
            )
            self.assertEqual(len(posts), 1, "Reload must not replay the last command")
            context.close()
            browser.close()

    def test_report_menu_navigation_mobile_and_reduced_motion(self):
        from playwright.sync_api import sync_playwright, expect

        with isolated_server() as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(channel="chrome", headless=True)
            context = browser.new_context(viewport=dict(width=1440, height=1000))
            page = context.new_page()
            errors = []
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.goto(url + "/reports")
            expect(page.locator("[data-tool-card]")).to_have_count(3)
            expect(page.locator("#source")).to_have_value("auto")
            expect(page.locator("#start")).to_be_disabled()
            self.assertEqual(
                page.locator("body").evaluate(
                    "node=>getComputedStyle(node).backgroundColor"
                ),
                "rgb(243, 242, 238)",
            )
            page.get_by_role("button", name="Mở menu", exact=True).click()
            expect(page.locator("#workspace-menu")).to_be_visible()
            expect(page.locator("#site-menu-open")).to_have_attribute(
                "aria-expanded", "true"
            )
            capture(page, "nothing-menu.png", full_page=False)
            page.keyboard.press("Escape")
            expect(page.locator("#workspace-menu")).not_to_be_visible()
            expect(page.locator("#site-menu-open")).to_be_focused()
            page.get_by_role("button", name="Mở menu", exact=True).click()
            page.locator('#workspace-menu a[href="/reports?tool=hfpa"]').click()
            expect(page.locator("#tool-title")).to_have_text("HFPA")
            expect(page.locator("#source")).to_have_value("auto")
            expect(page.locator("#groups .group")).to_have_count(2)
            page.locator('[data-tool="reinspection"]').click()
            expect(page.locator("#tool-title")).to_have_text("Re-Inspection")
            page.evaluate("window.scrollTo(0,0)")
            capture(page, "nothing-reports-desktop.png")
            page.set_viewport_size(dict(width=1440, height=1200))
            capture(page, "nothing-reports-preview.png", full_page=False)
            page.set_viewport_size(dict(width=390, height=844))
            page.evaluate("window.scrollTo(0,0)")
            self.assertLessEqual(
                page.evaluate("document.documentElement.scrollWidth"), 390
            )
            capture(page, "nothing-reports-mobile.png")
            page.get_by_role("button", name="Mở menu", exact=True).click()
            self.assertLessEqual(
                page.locator("#workspace-menu").bounding_box()["width"], 390
            )
            page.locator("#site-menu-close").click()
            self.assertEqual(errors, [])
            context.close()
            reduced = browser.new_context(reduced_motion="reduce")
            page = reduced.new_page()
            page.goto(url + "/reports?tool=copq")
            expect(page.locator("#tool-title")).to_have_text("combine_COPQ")
            duration = page.locator(".reports-hero").evaluate(
                "node=>parseFloat(getComputedStyle(node).animationDuration)"
            )
            self.assertLess(duration, 0.02)
            page.locator('[data-tool="hfpa"]').click()
            expect(page.locator("#tool-title")).to_have_text("HFPA")
            self.assertEqual(
                page.locator("#workspace").evaluate(
                    "node=>node.getAnimations().length"
                ),
                0,
            )
            reduced.close()
            browser.close()

    def test_dashboard_controls_modal_theme_and_small_screen(self):
        from playwright.sync_api import sync_playwright, expect

        with isolated_server() as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(channel="chrome", headless=True)
            page = browser.new_page(viewport=dict(width=1440, height=1100))
            errors = []
            posts = []
            page.on("pageerror", lambda error: errors.append(str(error)))

            def api(route):
                request = route.request
                if request.method != "GET":
                    posts.append((request.url, request.post_data_json))
                path = request.url.split("/api/", 1)[-1]
                data = {"success": True}
                if path.startswith("status"):
                    data = {"connected": False}
                elif path == "memory":
                    data = {"modules": [], "qms_modules": []}
                elif path == "factories":
                    data = {"factories": []}
                elif path == "nav/config":
                    data = {"success": True, "tiles": server.DEFAULT_NAV_TILES}
                elif path == "page/inspect":
                    data = {
                        "success": False,
                        "error": "Chưa kết nối SFC. Đăng nhập để tiếp tục.",
                    }
                elif path == "switch_factory":
                    plants = request.post_data_json["factories"]
                    data = {
                        "success": True,
                        "selected_plants": plants,
                        "current_plant": plants[-1],
                    }
                route.fulfill(json=data)

            page.route("**/api/**", api)
            external = []
            page.on(
                "request",
                lambda request: external.append(request.url)
                if not request.url.startswith(url)
                else None,
            )
            page.goto(url)
            expect(page.locator("#workspace-heading")).to_be_visible()
            expect(page.locator("#cdp-text")).to_have_text("SFC: Chưa kết nối")
            expect(page.locator("body")).to_have_attribute("data-theme", "light")
            page.locator("#plant-btn-VH2").click()
            expect(page.locator("#plant-btn-VH2")).to_have_class("plant-btn active")
            self.assertEqual(posts[-1][1]["factories"], ["VH", "VH2"])
            page.locator("#btn-customize-nav").click()
            expect(page.locator("#nav-modal-backdrop")).to_be_visible()
            page.locator("#nav-modal-backdrop .modal-close-btn").click()
            expect(page.locator("#nav-modal-backdrop")).not_to_be_visible()
            page.locator("#theme-toggle").click()
            expect(page.locator("body")).to_have_attribute("data-theme", "dark")
            page.locator("#theme-toggle").click()
            expect(page.locator("body")).to_have_attribute("data-theme", "light")
            expect(page.locator("#toast")).not_to_have_class("show", timeout=6000)
            page.evaluate("window.scrollTo(0,0)")
            capture(page, "nothing-dashboard-desktop.png")
            page.set_viewport_size(dict(width=390, height=844))
            self.assertLessEqual(
                page.evaluate("document.documentElement.scrollWidth"), 390
            )
            capture(page, "nothing-dashboard-mobile.png")
            self.assertEqual(external, [])
            self.assertEqual(errors, [])
            browser.close()


if __name__ == "__main__":
    unittest.main()
