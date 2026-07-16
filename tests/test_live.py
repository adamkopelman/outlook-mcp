"""Live tests against a real, running Outlook. Skipped by default — set
OUTLOOK_MCP_LIVE_TESTS=1 to run them (mirrors outlook-mcp-rs's
`cargo test -- --ignored` opt-in convention, adapted to pytest's skip
marker since Python has no built-in `--ignored`-equivalent test tag).

Run explicitly:
    OUTLOOK_MCP_LIVE_TESTS=1 pytest tests/test_live.py -v
"""

import os

import pytest

from outlook_mcp.outlook.client import WindowsOutlookClient

pytestmark = pytest.mark.skipif(
    os.environ.get("OUTLOOK_MCP_LIVE_TESTS") != "1",
    reason="set OUTLOOK_MCP_LIVE_TESTS=1 to run live Outlook tests",
)


@pytest.fixture
def client():
    return WindowsOutlookClient()


def test_categories_round_trip_on_a_real_task(client):
    created = client.create_task("outlook-mcp P1-python live categories probe")
    task_id = created["id"]
    try:
        from outlook_mcp.outlook.client import _get_item_categories, _set_item_categories
        _, ns = client._mapi()
        item = client._get_item(ns, task_id)
        _set_item_categories(item, ["Red Category", "Blue Category"])
        item.Save()
        cats = _get_item_categories(item)
        assert set(cats) == {"Red Category", "Blue Category"}

        tasks = client.list_tasks(include_completed=True)
        found = next((t for t in tasks if t["id"] == task_id), None)
        assert found is not None
        assert set(found["categories"]) == {"Red Category", "Blue Category"}
        assert found["status"] == "not_started"
        assert found["importance"] == "normal"
    finally:
        _, ns = client._mapi()
        item = client._get_item(ns, task_id)
        item.Delete()
