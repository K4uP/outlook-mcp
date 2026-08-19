"""Outlook MCP Server — exposes Outlook operations as MCP tools via stdio."""

import json
import sys
import threading
from typing import Optional

import anyio
from mcp.server.fastmcp import FastMCP

from . import outlook

mcp = FastMCP("outlook-mcp-server")

_com_lock = threading.Lock()


def _json(data) -> str:
    """Format response as indented JSON string."""
    return json.dumps(data, ensure_ascii=False, indent=2, default=str)


def _run_com(func, *args, **kwargs):
    import pythoncom
    with _com_lock:
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()


# ─── 1. Accounts & Folders ─────────────────────────────────────


@mcp.tool()
async def listAccounts() -> str:
    """List all configured email accounts in Outlook (name, email, type)."""
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.list_accounts))
    return _json(result)


@mcp.tool()
async def listFolders(accountId: Optional[str] = None) -> str:
    """List all mail folders with message counts.

    Args:
        accountId: Optional account name to filter by.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.list_folders, accountId))
    return _json(result)


@mcp.tool()
async def createFolder(parentFolderPath: str, name: str) -> str:
    """Create a new subfolder under the specified parent folder.

    Args:
        parentFolderPath: Path of the parent folder (e.g. "AccountName/Inbox").
        name: Name for the new folder.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.create_folder, parentFolderPath, name))
    return _json(result)


# ─── 2. Search & Read ──────────────────────────────────────────


@mcp.tool()
async def searchMessages(
    query: str,
    startDate: Optional[str] = None,
    endDate: Optional[str] = None,
    maxResults: int = 50,
    sortOrder: str = "desc",
) -> str:
    """Search messages by keyword across all folders.

    Args:
        query: Search keyword (matches subject, sender, recipient).
        startDate: Start date filter (ISO 8601).
        endDate: End date filter (ISO 8601).
        maxResults: Maximum results to return (default 50, max 200).
        sortOrder: "desc" (newest first, default) or "asc".
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.search_messages, query, startDate, endDate, maxResults, sortOrder))
    return _json(result)


@mcp.tool()
async def getRecentMessages(
    folderPath: Optional[str] = None,
    daysBack: int = 7,
    maxResults: int = 50,
    unreadOnly: bool = False,
) -> str:
    """Get recent messages, optionally filtered by folder and read status.

    Args:
        folderPath: Folder path (default: all inboxes).
        daysBack: Number of days to look back (default 7).
        maxResults: Maximum results (default 50, max 200).
        unreadOnly: Only return unread messages (default false).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.get_recent_messages, folderPath, daysBack, maxResults, unreadOnly))
    return _json(result)


@mcp.tool()
async def getMessage(
    messageId: str,
    folderPath: str,
    saveAttachments: bool = False,
) -> str:
    """Read full message content including body and attachments.

    Args:
        messageId: The message EntryID.
        folderPath: Folder path where the message resides.
        saveAttachments: Whether to save attachments to a temp directory (default false).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.get_message, messageId, folderPath, saveAttachments))
    return _json(result)


# ─── 3. Mail Operations ────────────────────────────────────────


@mcp.tool()
async def sendMail(
    to: str,
    subject: str,
    body: str,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    isHtml: bool = False,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Compose a new email and open it in Outlook for user confirmation.

    Args:
        to: Recipient email address(es).
        subject: Email subject.
        body: Email body text.
        cc: CC recipients.
        bcc: BCC recipients.
        isHtml: Whether body is HTML (default false).
        fromAccount: Send-from account (email or display name).
        attachments: List of file paths to attach.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.send_mail, to, subject, body, cc, bcc, isHtml, fromAccount, attachments))
    return _json(result)


@mcp.tool()
async def replyToMessage(
    messageId: str,
    folderPath: str,
    body: str,
    replyAll: bool = False,
    isHtml: bool = False,
    to: Optional[str] = None,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Reply to a message (opens compose window for user confirmation).

    Args:
        messageId: Original message EntryID.
        folderPath: Folder path of the original message.
        body: Reply body text.
        replyAll: Reply to all recipients (default false).
        isHtml: Whether body is HTML.
        to: Override recipient.
        cc: CC recipients.
        bcc: BCC recipients.
        fromAccount: Send-from account.
        attachments: File paths to attach.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(
        outlook.reply_to_message,
        messageId, folderPath, body, replyAll, isHtml, to, cc, bcc, fromAccount, attachments
    ))
    return _json(result)


@mcp.tool()
async def forwardMessage(
    messageId: str,
    folderPath: str,
    to: str,
    body: Optional[str] = None,
    isHtml: bool = False,
    cc: Optional[str] = None,
    bcc: Optional[str] = None,
    fromAccount: Optional[str] = None,
    attachments: Optional[list[str]] = None,
) -> str:
    """Forward a message (opens compose window for user confirmation).

    Args:
        messageId: Original message EntryID.
        folderPath: Folder path of the original message.
        to: Forward recipient.
        body: Additional body text prepended to the original.
        isHtml: Whether body is HTML.
        cc: CC recipients.
        bcc: BCC recipients.
        fromAccount: Send-from account.
        attachments: Additional file paths to attach.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(
        outlook.forward_message,
        messageId, folderPath, to, body, isHtml, cc, bcc, fromAccount, attachments
    ))
    return _json(result)


@mcp.tool()
async def updateMessage(
    messageId: str,
    folderPath: str,
    read: Optional[bool] = None,
    flagged: Optional[bool] = None,
    moveTo: Optional[str] = None,
    trash: bool = False,
) -> str:
    """Update message status: mark read/unread, flag/unflag, move, or trash.

    Args:
        messageId: Message EntryID.
        folderPath: Current folder path.
        read: Set read status (true=read, false=unread).
        flagged: Set flag status (true=flagged, false=unflagged).
        moveTo: Target folder path to move message to (cannot use with trash).
        trash: Move to deleted items (cannot use with moveTo).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.update_message, messageId, folderPath, read, flagged, moveTo, trash))
    return _json(result)


@mcp.tool()
async def deleteMessages(
    messageIds: list[str],
    folderPath: str,
) -> str:
    """Batch delete messages.

    Args:
        messageIds: List of message EntryIDs to delete.
        folderPath: Folder path where the messages reside.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.delete_messages, messageIds, folderPath))
    return _json(result)


# ─── 4. Contacts ───────────────────────────────────────────────


@mcp.tool()
async def searchContacts(query: str) -> str:
    """Search contacts by name or email.

    Args:
        query: Search keyword (matches name or email).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.search_contacts, query))
    return _json(result)


# ─── 5. Calendar ───────────────────────────────────────────────


@mcp.tool()
async def listCalendars() -> str:
    """List all calendars in Outlook."""
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.list_calendars))
    return _json(result)


@mcp.tool()
async def createEvent(
    title: str,
    startDate: str,
    endDate: Optional[str] = None,
    location: Optional[str] = None,
    description: Optional[str] = None,
    calendarId: Optional[str] = None,
    allDay: bool = False,
) -> str:
    """Create a calendar event (opens in Outlook for user confirmation).

    Args:
        title: Event title.
        startDate: Start time (ISO 8601).
        endDate: End time (default: start + 1 hour).
        location: Event location.
        description: Event description.
        calendarId: Target calendar ID/path.
        allDay: Whether this is an all-day event (default false).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(
        outlook.create_event,
        title, startDate, endDate, location, description, calendarId, allDay
    ))
    return _json(result)


@mcp.tool()
async def listEvents(
    calendarPath: Optional[str] = None,
    startDate: Optional[str] = None,
    endDate: Optional[str] = None,
    maxResults: int = 50,
) -> str:
    """List calendar events in a date range. Without arguments returns next 7 days.

    Args:
        calendarPath: Calendar folder path (default: all calendars).
        startDate: Start of range, ISO 8601 (default: now).
        endDate: End of range, ISO 8601 (default: now + 7 days).
        maxResults: Maximum events to return (default 50, max 500).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.list_events, calendarPath, startDate, endDate, maxResults))
    return _json(result)


@mcp.tool()
async def getEvent(eventId: str, calendarPath: str) -> str:
    """Get a single calendar event by its EntryID.

    Args:
        eventId: The event EntryID.
        calendarPath: Calendar folder path where the event lives.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.get_event, eventId, calendarPath))
    return _json(result)


@mcp.tool()
async def updateEvent(
    eventId: str,
    calendarPath: str,
    subject: Optional[str] = None,
    startDate: Optional[str] = None,
    endDate: Optional[str] = None,
    location: Optional[str] = None,
    description: Optional[str] = None,
) -> str:
    """Update a calendar event (opens in Outlook for user confirmation).

    Args:
        eventId: The event EntryID.
        calendarPath: Calendar folder path where the event lives.
        subject: New title (optional).
        startDate: New start time, ISO 8601 (optional).
        endDate: New end time, ISO 8601 (optional).
        location: New location (optional).
        description: New description (optional).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(
        outlook.update_event,
        eventId, calendarPath, subject, startDate, endDate, location, description
    ))
    return _json(result)


@mcp.tool()
async def deleteEvent(eventId: str, calendarPath: str) -> str:
    """Delete a calendar event.

    Args:
        eventId: The event EntryID.
        calendarPath: Calendar folder path where the event lives.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.delete_event, eventId, calendarPath))
    return _json(result)


# ─── 6. Tasks ──────────────────────────────────────────────────


@mcp.tool()
async def listTasks(
    includeCompleted: bool = False,
    maxResults: int = 100,
) -> str:
    """List tasks from the Outlook Tasks folder.

    Args:
        includeCompleted: Include completed tasks (default false).
        maxResults: Maximum tasks to return (default 100, max 500).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.list_tasks, includeCompleted, maxResults))
    return _json(result)


@mcp.tool()
async def createTask(
    subject: str,
    dueDate: Optional[str] = None,
    startDate: Optional[str] = None,
    body: Optional[str] = None,
    priority: int = 1,
) -> str:
    """Create a new task (opens in Outlook for user confirmation).

    Args:
        subject: Task title.
        dueDate: Due date, ISO 8601 (optional).
        startDate: Start date, ISO 8601 (optional).
        body: Task notes (optional).
        priority: 0=low, 1=normal (default), 2=high.
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(outlook.create_task, subject, dueDate, startDate, body, priority))
    return _json(result)


@mcp.tool()
async def updateTask(
    taskId: str,
    subject: Optional[str] = None,
    dueDate: Optional[str] = None,
    status: Optional[int] = None,
    percentComplete: Optional[int] = None,
    body: Optional[str] = None,
) -> str:
    """Update an existing task (opens in Outlook for user confirmation).

    Args:
        taskId: Task EntryID.
        subject: New title (optional).
        dueDate: New due date, ISO 8601 (optional).
        status: 0=notStarted, 1=inProgress, 2=complete, 3=waitingOnOthers, 4=deferred (optional).
        percentComplete: 0-100 (optional).
        body: New notes (optional).
    """
    result = await anyio.to_thread.run_sync(lambda: _run_com(
        outlook.update_task,
        taskId, subject, dueDate, status, percentComplete, body
    ))
    return _json(result)


# ─── Entry point ────────────────────────────────────────────────


def main():
    """Run the MCP server with stdio transport."""
    if "--version" in sys.argv or "-V" in sys.argv:
        try:
            from importlib.metadata import version
            ver = version("outlook-mcp-server")
        except Exception:
            ver = "dev"
        print(f"outlook-mcp-server {ver}")
        return
    if "--help" in sys.argv or "-h" in sys.argv:
        print("Usage: outlook-mcp-server")
        print("  MCP server for Outlook (stdio transport)")
        print()
        print("Options:")
        print("  -V, --version  Show version and exit")
        print("  -h, --help     Show this help and exit")
        return
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
