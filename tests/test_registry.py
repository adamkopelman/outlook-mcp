import asyncio

from outlook_mcp import server

EXPECTED_TOOLS = {
    # Email
    "list_folders", "list_emails", "search_emails", "get_email",
    "send_email", "create_draft", "reply_email", "move_email", "delete_email",
    # Calendar
    "list_events", "get_event", "create_event", "respond_to_meeting",
    # Attachments
    "list_attachments", "save_attachments",
    # Tasks
    "list_tasks", "create_task", "complete_task",
    # Notes
    "list_notes", "get_note", "create_note",
}


def list_tools():
    return asyncio.run(server.mcp.list_tools())


def test_all_expected_tools_registered():
    names = {tool.name for tool in list_tools()}
    assert names == EXPECTED_TOOLS


def test_every_tool_has_description_and_object_schema():
    for tool in list_tools():
        assert tool.description and tool.description.strip(), tool.name
        assert tool.inputSchema["type"] == "object", tool.name


def test_required_arguments_marked_required():
    schemas = {tool.name: tool.inputSchema for tool in list_tools()}
    assert set(schemas["send_email"]["required"]) == {"to", "subject", "body"}
    assert set(schemas["search_emails"]["required"]) == {"query"}
    assert set(schemas["create_event"]["required"]) == {"subject", "start", "end"}
    assert set(schemas["respond_to_meeting"]["required"]) == {"event_id", "response"}
    assert "required" not in schemas["list_folders"] or not schemas["list_folders"]["required"]
