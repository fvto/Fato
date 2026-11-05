"""Optional real browser test. CHINGLUH_UI_TESTS=1 uses a separate headless Chrome."""

import os
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sfc_report_tools as reports
import sfc_dashboard_server as server
from test_report_reliability import workbook
from werkzeug.serving import make_server


@unittest.skipUnless(
    os.environ.get("CHINGLUH_UI_TESTS") == "1",
    "Set CHINGLUH_UI_TESTS=1 for headless Chrome",
)
class BrowserTests(unittest.TestCase):
    def test_warning_button_toggles_brief_on_desktop_and_mobile(self):
        from playwright.sync_api import sync_playwright, expect
        from test_report_brief import fixture_brief

        run = dict(id="a" * 32, title="combine_COPQ", period="2026-04", status="done",
                   stage="Đã hoàn tất", started_at="2026-10-27T09:12:01+07:00", files=[],
                   requires_ack=True, export_review=dict(brief=fixture_brief(), checks=[], scope="Kiểm tra kết quả"))
        # All text from report data must stay text, including names and unknown labels.
        run["export_review"]["brief"]["issues"][0]["file"] = '<img src=x onerror="window.briefInjected=true">.xlsx'
        run["export_review"]["checks"] = [
            dict(rule="copq.row_count", status="passed", message="Số bản ghi khớp"),
            dict(rule="copq.quality", status="warning", message="Kiểm tra dữ liệu COPQ", mismatches=238,
                 issues=run["export_review"]["brief"]["issues"]),
        ]
        with (
            tempfile.TemporaryDirectory() as temporary,
            patch.object(reports, "RUN_ROOT", Path(temporary)),
            patch.object(reports, "runs", return_value=[run]),
            patch.object(reports, "registry", return_value=reports.registry(include_inputs=False)),
        ):
            http = make_server("127.0.0.1", 0, server.app, threaded=True)
            thread = threading.Thread(target=http.serve_forever, daemon=True)
            thread.start()
            try:
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch(channel="chrome", headless=True)
                    for width in (1280, 390):
                        page = browser.new_page(viewport=dict(width=width, height=1000))
                        errors = []
                        page.on("pageerror", lambda error: errors.append(str(error)))
                        page.goto(f"http://127.0.0.1:{http.server_port}/reports")
                        brief = page.locator(".run .report-brief")
                        expect(brief).to_be_hidden()
                        page.locator(".run > details > summary").click()
                        trigger = page.locator(".run [data-brief]")
                        expect(trigger).to_have_count(1)
                        expect(trigger).to_have_attribute("aria-expanded", "false")
                        trigger.click()
                        expect(brief).to_be_visible()
                        expect(trigger).to_have_attribute("aria-expanded", "true")
                        expect(brief.locator(".brief-list > li")).to_have_count(4)
                        expect(brief).to_contain_text("238 bản ghi cần xem lại")
                        expect(brief).to_contain_text("234 dòng · Thiếu Factory")
                        expect(brief).to_contain_text("JV: 234")
                        expect(brief).to_contain_text("722, 766")
                        expect(brief).to_contain_text("Giờ làm: 8")
                        expect(brief).to_contain_text("Số đôi: 95")
                        expect(brief).to_contain_text("có thể hợp lệ")
                        self.assertIsNone(page.evaluate("window.briefInjected"))
                        self.assertEqual(brief.locator("img").count(), 0)
                        self.assertLessEqual(page.evaluate("document.documentElement.scrollWidth"), width)
                        bounds = brief.bounding_box()
                        self.assertGreaterEqual(bounds["x"], 0)
                        self.assertLessEqual(bounds["x"] + bounds["width"], width)
                        if os.environ.get("CHINGLUH_DESIGN_SCREENSHOTS"):
                            screenshots = Path(os.environ["CHINGLUH_DESIGN_SCREENSHOTS"])
                            screenshots.mkdir(parents=True, exist_ok=True)
                            page.screenshot(path=str(screenshots / f"brief-{width}.png"))
                        page.keyboard.press("Escape")
                        expect(brief).to_be_hidden()
                        expect(trigger).to_be_focused()
                        trigger.click()
                        expect(brief).to_be_visible()
                        trigger.click()
                        expect(brief).to_be_hidden()
                        trigger.click()
                        expect(brief).to_be_visible()
                        page.locator("#results-heading").click()
                        expect(brief).to_be_hidden()
                        expect(page.locator("[data-ack]")).to_be_visible()
                        expect(page.locator("[data-ack]")).not_to_be_checked()
                        self.assertEqual(errors, [])
                        page.close()
                    browser.close()
            finally:
                http.shutdown()
                thread.join(timeout=5)

    def test_switch_tools_from_card_body_title_and_keyboard(self):
        from playwright.sync_api import sync_playwright, expect

        definitions = reports.registry(include_inputs=False)
        with (
            tempfile.TemporaryDirectory() as temporary,
            patch.object(reports, "RUN_ROOT", Path(temporary)),
            patch.object(reports, "registry", return_value=definitions),
            patch.object(reports.subprocess, "Popen") as spawn,
        ):
            http = make_server("127.0.0.1", 0, server.app, threaded=True)
            thread = threading.Thread(target=http.serve_forever, daemon=True)
            thread.start()
            try:
                with sync_playwright() as playwright:
                    browser = playwright.chromium.launch(
                        channel="chrome", headless=True
                    )
                    for legacy in (False, True):
                        with self.subTest(legacy=legacy):
                            context = browser.new_context()
                            page = context.new_page()
                            errors = []
                            page.on(
                                "pageerror", lambda error: errors.append(str(error))
                            )
                            if legacy:
                                tools = [
                                    {
                                        key: value
                                        for key, value in tool.items()
                                        if key
                                        not in (
                                            "automatic",
                                            "contract_version",
                                            "required_groups",
                                        )
                                    }
                                    for tool in definitions
                                ]
                                page.route(
                                    "**/api/report-tools?catalog=0",
                                    lambda route: route.fulfill(
                                        json=dict(tools=tools, runs=[])
                                    ),
                                )
                            page.goto(f"http://127.0.0.1:{http.server_port}/reports")
                            expect(page.locator("#tool-title")).to_have_text(
                                "Re-Inspection"
                            )
                            # The whole visible card used to be clickable, including its description.
                            for tool_id, title, mode_count, group_count in [
                                ("copq", "combine_COPQ", 3, 2),
                                ("hfpa", "HFPA", 3, 2),
                                ("reinspection", "Re-Inspection", 4, 3),
                                ("copq", "combine_COPQ", 3, 2),
                            ]:
                                card = page.locator(".tool").filter(
                                    has=page.locator(f'[data-tool="{tool_id}"]')
                                )
                                card.locator("p").click()
                                expect(page.locator("#tool-title")).to_have_text(title)
                                expect(page.locator("#mode option")).to_have_count(
                                    mode_count
                                )
                                expect(page.locator("#source")).to_have_value("auto")
                                expect(page.locator("#groups .group")).to_have_count(
                                    group_count
                                )
                            page.locator(".tool").filter(
                                has=page.locator('[data-tool="reinspection"]')
                            ).click(position=dict(x=12, y=12))
                            expect(page.locator("#tool-title")).to_have_text(
                                "Re-Inspection"
                            )
                            page.locator('[data-tool="hfpa"]').click()
                            expect(page.locator("#tool-title")).to_have_text("HFPA")
                            page.locator('[data-tool-card="copq"] .tool-object').click()
                            expect(page.locator("#tool-title")).to_have_text(
                                "combine_COPQ"
                            )
                            card = page.locator(".tool").filter(
                                has=page.locator('[data-tool="reinspection"]')
                            )
                            card.focus()
                            card.press("Enter")
                            expect(page.locator("#tool-title")).to_have_text(
                                "Re-Inspection"
                            )
                            card = page.locator(".tool").filter(
                                has=page.locator('[data-tool="copq"]')
                            )
                            card.focus()
                            card.press("Space")
                            expect(page.locator("#tool-title")).to_have_text(
                                "combine_COPQ"
                            )
                            page.locator('[data-favorite="hfpa"]').click()
                            expect(page.locator("#tool-title")).to_have_text(
                                "combine_COPQ"
                            )
                            self.assertEqual(errors, [])
                            context.close()
                    spawn.assert_not_called()
                    browser.close()
            finally:
                http.shutdown()
                thread.join(timeout=5)

    def test_upload_validate_execute_review_presets_cancel_and_download(self):
        from playwright.sync_api import sync_playwright, expect

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = workbook(root / "COPQ-VH.xlsx")
            (root / "engines/copq").mkdir(parents=True)
            definition = reports.TOOLS["copq"]
            tool = dict(
                id="copq",
                title=definition["title"],
                description=definition["description"],
                plants=definition["plants"],
                groups=[
                    dict(id=g, title=v[0], formats=reports.input_formats("copq", g))
                    for g, v in definition["groups"].items()
                ],
                modes=[dict(id=k, title=v) for k, v in definition["modes"].items()],
                default_period="2026-09",
                ready=True,
                missing_dependencies=[],
                inputs=[],
                automatic=True,
                contract_version=1,
                required_groups=definition["required_groups"],
                max_file_bytes=52428800,
            )
            with (
                patch.object(reports, "RUN_ROOT", root / "runs"),
                patch.object(reports, "TOOLS_ROOT", root / "engines"),
                patch.object(reports, "registry", return_value=[tool]),
                patch.object(reports, "input_catalog", return_value=[]),
                patch.object(reports.subprocess, "Popen") as spawn,
            ):
                http = make_server("127.0.0.1", 0, server.app, threaded=True)
                thread = threading.Thread(target=http.serve_forever, daemon=True)
                thread.start()
                try:
                    with sync_playwright() as playwright:
                        browser = playwright.chromium.launch(
                            channel="chrome", headless=True
                        )
                        page = browser.new_page(viewport=dict(width=1280, height=1000))
                        errors = []
                        page.on("pageerror", lambda error: errors.append(str(error)))
                        page.goto(f"http://127.0.0.1:{http.server_port}/reports")
                        expect(page.locator("#tool-title")).to_have_text("combine_COPQ")
                        page.locator("#mode").select_option("combine")
                        page.locator("#source").select_option("upload")
                        page.locator("#save-preset").click()
                        page.locator("#source").select_option("local")
                        page.locator("#load-preset").click()
                        expect(page.locator("#source")).to_have_value("upload")
                        page.locator("[data-favorite]").click()
                        expect(page.locator("[data-favorite]")).to_have_attribute(
                            "aria-pressed", "true"
                        )
                        page.locator("#search").fill("missing-tool")
                        expect(page.locator("[data-tool]")).to_have_count(0)
                        page.locator("#search").fill("")
                        expect(page.locator("[data-tool]")).to_have_count(1)
                        page.locator('[data-upload="copq"]').set_input_files(
                            str(source)
                        )
                        expect(page.locator("#start")).to_be_disabled()
                        page.locator("#validate").click()
                        expect(page.locator("#start")).to_be_enabled(timeout=15000)
                        spawn.assert_not_called()
                        page.locator("#start").click()
                        expect(page.locator(".run")).to_have_count(1)
                        spawn.assert_called_once()
                        # Refresh never resubmits. Cancellation is a marker, not process termination.
                        page.locator("#refresh").click()
                        expect(page.locator("[data-cancel]")).to_be_visible()
                        page.locator("[data-cancel]").click()
                        expect(
                            page.get_by_text("Đã yêu cầu hủy.", exact=False)
                        ).to_be_visible()
                        run_id = next((root / "runs").glob("*/run.json")).parent.name
                        run = reports.read_json(root / "runs" / run_id / "run.json")
                        output = root / "runs" / run_id / "workspace/Output"
                        output.mkdir()
                        workbook(output / "result.xlsx")
                        for step in run["steps"]:
                            step.update(
                                status="done",
                                started_at=run["started_at"],
                                finished_at=reports.now(),
                            )
                        run.update(
                            status="done",
                            stage="Đã hoàn tất; xem đối chiếu trước khi tải.",
                            files=reports.collect_files(output.parent.parent),
                            requires_ack=True,
                            export_review=dict(
                                checks=[
                                    dict(
                                        rule="fixture",
                                        status="warning",
                                        message="Kiểm tra mẫu: sai lệch số dòng",
                                        expected=2,
                                        actual=1,
                                    )
                                ],
                                scope="Dữ liệu giả lập dùng để kiểm tra giao diện.",
                            ),
                        )
                        reports.write_json(root / "runs" / run_id / "run.json", run)
                        page.locator("#refresh").click()
                        page.locator(".run > details > summary").click()
                        link = page.get_by_role("link", name="result.xlsx")
                        link.click()
                        expect(page.locator("#notice")).to_contain_text(
                            "đánh dấu xác nhận"
                        )
                        page.locator("[data-ack]").check()
                        with page.expect_download() as download_info:
                            link.click()
                        self.assertEqual(
                            download_info.value.suggested_filename, "result.xlsx"
                        )
                        spawn.assert_called_once()
                        self.assertEqual(errors, [])
                        screenshot = os.environ.get("CHINGLUH_TEST_SCREENSHOT")
                        if screenshot:
                            page.screenshot(path=screenshot, full_page=True)
                        page.set_viewport_size(dict(width=390, height=844))
                        self.assertLessEqual(
                            page.evaluate("document.documentElement.scrollWidth"), 390
                        )
                        # A server still running the old API must keep its automatic
                        # source and original acquire-missing route usable.
                        legacy_tool = {
                            key: value
                            for key, value in tool.items()
                            if key
                            not in ("automatic", "contract_version", "required_groups")
                        }
                        legacy_tool["inputs"] = [
                            dict(
                                id="local:copq:COPQ-VH.xlsx",
                                group="copq",
                                name="COPQ-VH.xlsx",
                                period="2026-09",
                                plant="VH",
                                source="local",
                                size=100,
                            ),
                            dict(
                                id="sfc:COPQ-VH.xlsx",
                                group="copq",
                                name="COPQ-VH.xlsx",
                                period="2026-10",
                                plant="VH",
                                source="sfc",
                                size=100,
                            ),
                        ]
                        page.route(
                            "**/api/report-tools?catalog=0",
                            lambda route: route.fulfill(
                                json=dict(tools=[legacy_tool], runs=[])
                            ),
                        )
                        legacy_requests = []
                        page.on(
                            "request",
                            lambda request: legacy_requests.append(request.url),
                        )
                        page.goto(f"http://127.0.0.1:{http.server_port}/reports")
                        expect(page.locator("#source")).to_have_value("auto")
                        expect(page.locator("#source option")).to_have_count(4)
                        self.assertFalse(
                            page.locator('#source option[value="auto"]').evaluate(
                                "node=>node.hidden"
                            )
                        )
                        expect(page.locator("#groups")).to_contain_text(
                            "COPQ Data · 7 nhà máy"
                        )
                        expect(page.locator("#start")).to_be_enabled()
                        expect(page.locator("#validate")).to_be_hidden()
                        page.locator("#mode").select_option("combine")
                        expect(page.locator("#groups")).not_to_contain_text(
                            "FTT Internal"
                        )
                        page.locator("#source").select_option("local")
                        expect(page.locator("#period")).to_have_value("2026-09")
                        expect(page.locator("[data-input]")).to_be_checked()
                        page.locator("#source").select_option("sfc")
                        expect(page.locator("#period")).to_have_value("2026-10")
                        expect(page.locator("[data-input]")).not_to_be_checked()
                        page.locator("#source").select_option("upload")
                        expect(page.locator('[data-upload="copq"]')).to_be_visible()
                        page.locator("#source").select_option("auto")
                        page.locator("#start").click()
                        expect(page.locator(".run")).to_have_count(1)
                        spawn.assert_called()
                        self.assertEqual(spawn.call_count, 2)
                        run_file = max(
                            (root / "runs").glob("*/run.json"),
                            key=lambda path: path.stat().st_mtime_ns,
                        )
                        auto_run = reports.read_json(run_file)
                        self.assertTrue(auto_run["acquire_missing"])
                        self.assertEqual(auto_run["acquisition"]["total"], 7)
                        self.assertEqual(len(auto_run["acquisition"]["missing"]), 7)
                        self.assertFalse(
                            any(
                                "/inputs?" in url or url.endswith("/validate")
                                for url in legacy_requests
                            )
                        )
                        self.assertEqual(errors, [])
                        browser.close()
                finally:
                    http.shutdown()
                    thread.join(timeout=5)


if __name__ == "__main__":
    unittest.main()
