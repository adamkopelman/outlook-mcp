"""Test harness for exercising WindowsOutlookClient without Outlook.

Importing this module:
  * installs minimal stand-ins for ``pythoncom``, ``pywintypes`` and
    ``win32com.client`` when they are not importable (i.e. off Windows), so
    ``outlook_mcp.outlook.client`` can be imported anywhere;
  * provides tiny fake COM objects (mail items, attachments, items
    collections, folders, namespace) that client tests can wire up by
    monkeypatching ``WindowsOutlookClient._mapi``.

On Windows the real pywin32 modules are used untouched; the fakes are plain
Python objects, so tests behave the same on every platform.

Fakes only model what tests need. ``FakeItems.Restrict`` cannot evaluate
JET/DASL filter strings: it records the filter and returns a collection
chosen by an optional ``restrict_hook(filter) -> list | None`` (``None``
means "keep the same items").
"""

import sys
import types

try:  # pragma: no cover - exercised implicitly per platform
    import pythoncom  # noqa: F401
    import pywintypes  # noqa: F401
    import win32com.client  # noqa: F401
except ImportError:
    class _ComError(Exception):
        """Stand-in for pywintypes.com_error: args are
        (hresult, strerror, excepinfo, argerror)."""

    _pythoncom = types.ModuleType("pythoncom")
    _pythoncom.CoInitialize = lambda: None
    _pythoncom.CoUninitialize = lambda: None

    _pywintypes = types.ModuleType("pywintypes")
    _pywintypes.com_error = _ComError

    def _dispatch(*_args, **_kwargs):
        raise _ComError(-2147221005, "Invalid class string", None, None)

    _win32com = types.ModuleType("win32com")
    _win32com_client = types.ModuleType("win32com.client")
    _win32com_client.Dispatch = _dispatch
    _win32com.client = _win32com_client

    sys.modules.setdefault("pythoncom", _pythoncom)
    sys.modules.setdefault("pywintypes", _pywintypes)
    sys.modules.setdefault("win32com", _win32com)
    sys.modules.setdefault("win32com.client", _win32com_client)

import pywintypes  # noqa: E402  (real or stub, whichever is installed)


def com_error(message="COM failure", hresult=-2147352567):
    """Build a com_error shaped like pywin32's (works with the stub too)."""
    return pywintypes.com_error(hresult, message, (0, "Outlook", message, None, 0, 0), None)


class FakePropertyAccessor:
    """MAPI PropertyAccessor backed by a dict keyed by schema name.
    Missing properties raise com_error, like Outlook does."""

    def __init__(self, props=None):
        self.props = dict(props or {})
        self.set_calls = []

    def GetProperty(self, schema_name):
        if schema_name not in self.props:
            raise com_error("The property is unknown or cannot be found.")
        return self.props[schema_name]

    def SetProperty(self, schema_name, value):
        self.set_calls.append((schema_name, value))
        self.props[schema_name] = value


class FakeAttachment:
    def __init__(self, filename="file.bin", data=b"", props=None, att_type=1):
        self.FileName = filename
        self.DisplayName = filename
        self.data = data
        self.Size = len(data)
        self.Type = att_type
        self.PropertyAccessor = FakePropertyAccessor(props)
        self.saved_to = []

    def SaveAsFile(self, path):
        with open(path, "wb") as fh:
            fh.write(self.data)
        self.saved_to.append(path)


class FakeAttachments:
    """1-based collection, like Outlook's Attachments."""

    def __init__(self, attachments=None):
        self._items = list(attachments or [])
        self.add_calls = []

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]

    def Add(self, source, att_type=1, position=None, display_name=None):
        self.add_calls.append((source, att_type, position, display_name))
        with open(source, "rb") as fh:
            data = fh.read()
        name = display_name or source.replace("\\", "/").rsplit("/", 1)[-1]
        att = FakeAttachment(name, data, att_type=att_type)
        self._items.append(att)
        return att

    def __iter__(self):
        return iter(list(self._items))


class FakeStore:
    def __init__(self, store_id="store-1", default_folders=None):
        self.StoreID = store_id
        self.default_folders = dict(default_folders or {})

    def GetDefaultFolder(self, folder_id):
        return self.default_folders[folder_id]


class FakeItems:
    """Items collection. Iterable, Count-able, records Restrict/Sort."""

    def __init__(self, items=None, restrict_hook=None, log=None):
        self._items = list(items or [])
        self.restrict_hook = restrict_hook
        self.log = log if log is not None else []

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]

    def Restrict(self, flt):
        self.log.append(("Restrict", flt))
        chosen = self.restrict_hook(flt, list(self._items)) if self.restrict_hook else None
        return FakeItems(self._items if chosen is None else chosen,
                         restrict_hook=self.restrict_hook, log=self.log)

    def Sort(self, prop, descending=False):
        self.log.append(("Sort", prop, descending))

    def __iter__(self):
        return iter(list(self._items))

    def remove(self, item):
        self._items.remove(item)

    def append(self, item):
        self._items.append(item)


class FakeFolder:
    def __init__(self, name="Inbox", entry_id=None, items=None, store=None,
                 folders=None, parent=None):
        self.Name = name
        self.EntryID = entry_id or f"folder-{name}"
        self.Items = items if isinstance(items, FakeItems) else FakeItems(items)
        self.Store = store or FakeStore()
        self.StoreID = self.Store.StoreID
        self.Folders = FakeItems(folders)
        self.Parent = parent
        self.deleted = False
        for item in self.Items:
            item.Parent = self

    def add(self, item):
        item.Parent = self
        self.Items.append(item)
        return item

    def Delete(self):
        self.deleted = True
        if self.Parent is not None:
            self.Parent.Folders.remove(self)


_id_counter = [0]


def _next_entry_id():
    _id_counter[0] += 1
    return f"entry-{_id_counter[0]}"


class FakeMail:
    """MailItem stand-in. Any keyword becomes an attribute."""

    def __init__(self, parent=None, attachments=None, **attrs):
        self.EntryID = attrs.pop("EntryID", None) or _next_entry_id()
        defaults = dict(Subject="", Body="", HTMLBody="", To="", CC="", BCC="",
                        SenderName="", SenderEmailAddress="", UnRead=False,
                        Categories="", MessageClass="IPM.Note", Sent=False,
                        BodyFormat=1, Importance=1, FlagStatus=0,
                        ReceivedTime=None)
        defaults.update(attrs)
        for k, v in defaults.items():
            setattr(self, k, v)
        self.Attachments = (attachments if isinstance(attachments, FakeAttachments)
                            else FakeAttachments(attachments))
        self.PropertyAccessor = FakePropertyAccessor()
        self.Parent = parent
        self.saved = 0
        self.sent = False
        self.deleted = False
        self.calls = []

    def Save(self):
        self.saved += 1

    def Send(self):
        self.sent = True

    def Delete(self):
        """Mirrors Outlook: outside Deleted Items this moves the item there
        (if the store has one wired up); inside it, deletes permanently."""
        self.calls.append("Delete")
        parent = self.Parent
        deleted_folder = None
        if parent is not None:
            deleted_folder = parent.Store.default_folders.get(3)
        if parent is not None:
            parent.Items.remove(self)
        if deleted_folder is not None and deleted_folder is not parent:
            self.EntryID = _next_entry_id()
            deleted_folder.add(self)
        else:
            self.deleted = True
            self.Parent = None

    def Move(self, folder):
        self.calls.append(("Move", folder.Name))
        if self.Parent is not None:
            self.Parent.Items.remove(self)
        self.EntryID = _next_entry_id()
        folder.add(self)
        return self


class FakeNamespace:
    """MAPI namespace: GetItemFromID over registered items, default folders."""

    def __init__(self, default_folders=None):
        self.default_folders = dict(default_folders or {})
        self._items = {}

    def register(self, item):
        self._items[item.EntryID] = item
        return f"{item.EntryID}|{item.Parent.StoreID if item.Parent else 'store-1'}"

    def GetItemFromID(self, entry_id, store_id=None):
        for item in self._items.values():
            if item.EntryID == entry_id:
                return item
        raise com_error("The operation failed. An object could not be found.")

    def GetDefaultFolder(self, folder_id):
        return self.default_folders[folder_id]


class FakeApp:
    def __init__(self, ns):
        self.ns = ns
        self.created = []

    def CreateItem(self, item_type):
        item = FakeMail()
        self.created.append((item_type, item))
        return item

    def GetNamespace(self, _name):
        return self.ns


def install(monkeypatch, ns, app=None):
    """Point WindowsOutlookClient._mapi at the given fake namespace and
    return a fresh client."""
    from outlook_mcp.outlook.client import WindowsOutlookClient

    app = app or FakeApp(ns)
    monkeypatch.setattr(WindowsOutlookClient, "_mapi", lambda self: (app, ns))
    return WindowsOutlookClient()
