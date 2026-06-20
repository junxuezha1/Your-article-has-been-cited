import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import app as app_module


class ProductBrandingTests(unittest.TestCase):
    def test_home_page_uses_product_identity_from_config(self):
        product = {
            "name": "{{MoreYield}}",
            "id": "{{King}}",
            "version": "{{VERSION}}",
            "logo": "assets/logo.png",
            "logo_light": "assets/logo-light.png",
            "favicon": "assets/favicon.ico",
            "app_icon": "assets/app.ico",
            "support": "{{my18874068595@gmail.com}}",
        }

        with patch.object(app_module, "load_config", return_value={"product": product}):
            html = app_module.app.test_client().get("/?journal=csu_social").get_data(as_text=True)

        self.assertIn("<title>{{MoreYield}} — 《中南大学学报（社会科学版）》</title>", html)
        self.assertIn('href="/assets/favicon.ico"', html)
        self.assertIn('src="/assets/logo-light.png"', html)
        self.assertIn("{{King}}", html)
        self.assertIn("{{VERSION}}", html)
        self.assertIn("{{my18874068595@gmail.com}}", html)

    def test_product_defaults_are_available_without_config_section(self):
        with patch.object(app_module, "load_config", return_value={}):
            html = app_module.app.test_client().get("/").get_data(as_text=True)

        self.assertIn("{{MoreYield}}", html)
        self.assertIn("{{King}}", html)
        self.assertIn("{{VERSION}}", html)
        self.assertIn("{{my18874068595@gmail.com}}", html)

    def test_provided_logo_assets_are_present_at_configured_paths(self):
        self.assertTrue((app_module.BASE_DIR / "assets" / "logo.png").exists())
        self.assertTrue((app_module.BASE_DIR / "assets" / "logo-light.png").exists())
        self.assertTrue((app_module.BASE_DIR / "assets" / "favicon.ico").exists())
        self.assertTrue((app_module.BASE_DIR / "assets" / "app.ico").exists())

    def test_configured_brand_assets_are_served(self):
        client = app_module.app.test_client()

        for asset_path in [
            "/assets/logo.png",
            "/assets/logo-light.png",
            "/assets/favicon.ico",
            "/assets/app.ico",
        ]:
            with self.subTest(asset_path=asset_path):
                response = client.get(asset_path, buffered=True)

                self.assertEqual(response.status_code, 200)
                self.assertGreater(len(response.get_data()), 0)
                response.close()

    def test_assets_route_serves_brand_files_from_project_assets(self):
        old_base = app_module.BASE_DIR
        try:
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                app_module.BASE_DIR = tmp_path
                assets_dir = tmp_path / "assets"
                assets_dir.mkdir()
                (assets_dir / "favicon.ico").write_bytes(b"ico")

                response = app_module.app.test_client().get("/assets/favicon.ico", buffered=True)
                response_data = response.get_data()

                self.assertEqual(response.status_code, 200)
                self.assertEqual(response_data, b"ico")
                response.close()
        finally:
            app_module.BASE_DIR = old_base


if __name__ == "__main__":
    unittest.main()
