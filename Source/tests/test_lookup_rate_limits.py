import unittest
from unittest.mock import patch
from urllib.error import HTTPError

from online_metadata import lookup_online_metadata


class RateLimitTests(unittest.TestCase):
    def test_limited_provider_is_not_retried_for_second_query(self):
        with patch("online_metadata.fetch_json", side_effect=HTTPError("test", 429, "limited", {}, None)) as fetch:
            with self.assertRaisesRegex(RuntimeError, "Google Books is temporarily limiting requests"):
                lookup_online_metadata({"isbn": "9781234567890", "title": "Example"}, service_filter="google")
            self.assertEqual(fetch.call_count, 1)
