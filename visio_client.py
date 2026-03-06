"""
Visio COM Automation Wrapper
Provides a clean Python interface to Microsoft Visio via win32com.
All styling follows STYLE_GUIDE.md conventions automatically.
"""

import json
import win32com.client
import os
import pythoncom
from pathlib import Path


# ── Azure stencil mapper (embedded) ─────────────────────────
STENCIL_MAP_FILE = Path(__file__).parent / "stencil_map.json"


def _get_stencil_search_roots(visio_app) -> list[Path]:
    """
    Build stencil search roots purely from Visio COM properties and the
    Windows Shell known-folder for Documents (which may differ from OneDrive).
    """
    roots: list[Path] = []

    # 1. Visio COM: MyShapesPath (user stencils)
    try:
        vpath = visio_app.MyShapesPath
        if vpath:
            roots.append(Path(vpath))
    except Exception:
        pass

    # 2. Visio COM: built-in content directory
    try:
        built_in = visio_app.GetBuiltInStencilFile(0, 0)  # any built-in
        if built_in:
            roots.append(Path(built_in).parent)
    except Exception:
        pass

    # 3. Visio COM: application install path
    try:
        roots.append(Path(visio_app.Path) / "Visio Content")
    except Exception:
        pass

    # 4. Windows Shell: real Documents folder (handles OneDrive redirect)
    try:
        import ctypes.wintypes
        CSIDL_PERSONAL = 5  # My Documents
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, 0, buf)
        if buf.value:
            roots.append(Path(buf.value) / "My Shapes")
    except Exception:
        pass

    # 5. Home fallback
    roots.append(Path.home() / "Documents" / "My Shapes")

    # Deduplicate while preserving priority order
    seen: set[str] = set()
    unique: list[Path] = []
    for r in roots:
        key = str(r).lower()
        if key not in seen:
            seen.add(key)
            unique.append(r)
    return unique


def _load_stencil_map() -> dict:
    """Load the Azure service → stencil mapping."""
    if STENCIL_MAP_FILE.exists():
        with open(STENCIL_MAP_FILE) as f:
            return json.load(f)
    return {}


def _fuzzy_key(name: str) -> str:
    """Normalize a service name to a stencil map key."""
    normalized = name.lower().replace(" ", "-").replace("_", "-").replace("azure-", "")
    if not normalized.startswith("azure/"):
        normalized = f"azure/{normalized}"
    return normalized


# ── Style constants from STYLE_GUIDE.md ─────────────────────
AZURE_COLORS = {
    "azure_blue":  (0, 120, 215),
    "dark_blue":   (0, 78, 152),
    "teal":        (0, 178, 148),
    "orange":      (255, 140, 0),
    "purple":      (135, 100, 184),
    "green":       (122, 184, 0),
    "red":         (232, 17, 35),
}

SHAPE_ROUNDING = 0.06          # inches
SHAPE_FILL_TRANSPARENCY = 15   # percent

CONNECTOR_END_ARROW = 4
CONNECTOR_ARROW_SIZE = 2
CONNECTOR_ROUNDING = 0.15      # inches
CONNECTOR_LINE_WEIGHT = 1      # pt
CONNECTOR_LABEL_FONT = 7       # pt
CONNECTOR_DASHED_PATTERN = 2

CONTAINER_LINE_PATTERN = 2     # dashed
CONTAINER_LINE_WEIGHT = 1      # pt
CONTAINER_FILL_TRANS = 60      # percent
CONTAINER_LABEL_PIN_Y = "Height*0.96"
CONTAINER_LABEL_FONT = 9       # pt

TIER_BAND_FILL_TRANS = 70      # percent
TIER_BAND_LABEL_FONT = 8      # pt


class VisioClient:
    """Wrapper around the Visio COM object model with STYLE_GUIDE.md compliance."""

    SHAPE_MASTERS = {
        "rectangle": ("BASIC_M.vssx", "Rectangle"),
        "square": ("BASIC_M.vssx", "Square"),
        "ellipse": ("BASIC_M.vssx", "Ellipse"),
        "circle": ("BASIC_M.vssx", "Circle"),
        "diamond": ("BASIC_M.vssx", "Diamond"),
        "triangle": ("BASIC_M.vssx", "Triangle"),
        "rounded_rectangle": ("BASIC_M.vssx", "Rounded rectangle"),
        "star": ("BASIC_M.vssx", "5-Point Star"),
    }

    def __init__(self):
        self._app = None
        self._stencils = {}
        self._stencil_map = _load_stencil_map()
        self._stencil_paths_registered = False

    @property
    def app(self):
        if self._app is None:
            pythoncom.CoInitialize()
            try:
                self._app = win32com.client.GetActiveObject("Visio.Application")
            except Exception:
                self._app = win32com.client.Dispatch("Visio.Application")
                self._app.Visible = True
        return self._app

    @property
    def active_doc(self):
        try:
            return self.app.ActiveDocument
        except Exception:
            return None

    @property
    def active_page(self):
        try:
            return self.app.ActivePage
        except Exception:
            return None

    # ── Stencil helpers ──────────────────────────────────────

    def _open_stencil(self, stencil_file: str):
        """Open a stencil file if not already loaded."""
        stencil_basename = os.path.basename(stencil_file).lower()

        # Check cache — but validate the COM reference is still alive
        if stencil_file in self._stencils:
            try:
                _ = self._stencils[stencil_file].Name
                return self._stencils[stencil_file]
            except Exception:
                del self._stencils[stencil_file]

        try:
            # Check if already open (compare by filename, not full path)
            for i in range(1, self.app.Documents.Count + 1):
                doc = self.app.Documents.Item(i)
                if doc.Name.lower() == stencil_basename:
                    self._stencils[stencil_file] = doc
                    return doc
            stencil = self.app.Documents.OpenEx(stencil_file, 4)  # visOpenDocked
            self._stencils[stencil_file] = stencil
            return stencil
        except Exception as e:
            raise RuntimeError(f"Cannot open stencil '{stencil_file}': {e}")

    def _register_stencil_paths(self):
        """Discover Azure stencil directories from Visio COM properties
        and register them in Visio's StencilPaths for filename resolution."""
        if self._stencil_paths_registered:
            return
        self._stencil_paths_registered = True

        roots = _get_stencil_search_roots(self.app)
        stencil_dirs: set[str] = set()

        for root in roots:
            if not root.exists():
                continue
            for vssx in root.rglob("Azure-*.vssx"):
                stencil_dirs.add(str(vssx.parent))

        if stencil_dirs:
            # Append to any existing StencilPaths
            existing = ""
            try:
                existing = self.app.StencilPaths or ""
            except Exception:
                pass
            all_paths = set(existing.split(";")) if existing else set()
            all_paths.update(stencil_dirs)
            all_paths.discard("")
            self.app.StencilPaths = ";".join(sorted(all_paths))

    def _open_azure_stencil(self, stencil_name: str):
        """Open an Azure stencil by logical name (e.g. 'Azure-Databases').

        Strategy (COM-first):
        1. Check if already open in Visio Documents collection
        2. Try Documents.OpenEx by filename — Visio searches its StencilPaths
        3. Fall back to explicit path from discovered directories
        """
        filename = f"{stencil_name}.vssx"

        # 1. Already open? (pure COM lookup)
        try:
            for i in range(1, self.app.Documents.Count + 1):
                doc = self.app.Documents.Item(i)
                if doc.Name.lower() == filename.lower():
                    self._stencils[filename] = doc
                    return doc
        except Exception:
            pass

        # 2. Register stencil paths from COM properties, then open by name
        self._register_stencil_paths()
        try:
            stencil = self.app.Documents.OpenEx(filename, 4)  # visOpenDocked
            self._stencils[filename] = stencil
            return stencil
        except Exception:
            pass

        # 3. Fallback: search roots for exact file and open by full path
        for root in _get_stencil_search_roots(self.app):
            if not root.exists():
                continue
            for vssx in root.rglob(filename):
                return self._open_stencil(str(vssx))

        raise RuntimeError(
            f"Azure stencil '{stencil_name}' not found. Ensure Azure Visio "
            f"stencils are installed in your My Shapes folder."
        )

    def _get_master(self, shape_type: str):
        """Get a Visio Master shape for the given type name."""
        shape_type_lower = shape_type.lower().replace(" ", "_")

        if shape_type_lower in self.SHAPE_MASTERS:
            stencil_file, master_name = self.SHAPE_MASTERS[shape_type_lower]
            stencil = self._open_stencil(stencil_file)
            return stencil.Masters.Item(master_name)

        # Try to find in any open stencil
        for i in range(1, self.app.Documents.Count + 1):
            doc = self.app.Documents.Item(i)
            try:
                return doc.Masters.Item(shape_type)
            except Exception:
                continue

        raise ValueError(
            f"Unknown shape type '{shape_type}'. Available: "
            f"{', '.join(self.SHAPE_MASTERS.keys())}"
        )

    def _resolve_azure_service(self, service: str):
        """Resolve an Azure service name to (stencil_name, master_name) or None."""
        key = _fuzzy_key(service)
        entry = self._stencil_map.get(key)
        if entry:
            return entry["stencil"], entry["master"]
        # Try exact match
        entry = self._stencil_map.get(service)
        if entry:
            return entry["stencil"], entry["master"]
        return None

    def _apply_shape_style(self, shape, fill_color: str | None = None):
        """Apply standard shape styling per STYLE_GUIDE.md."""
        try:
            shape.Cells("Rounding").FormulaU = f"{SHAPE_ROUNDING} in"
        except Exception:
            pass
        try:
            shape.Cells("FillForegndTrans").FormulaU = f"{SHAPE_FILL_TRANSPARENCY}%"
        except Exception:
            pass
        if fill_color:
            r, g, b = self._parse_color(fill_color)
            try:
                shape.Cells("FillForegnd").FormulaU = f"RGB({r},{g},{b})"
            except Exception:
                pass

    @staticmethod
    def _parse_color(color: str) -> tuple[int, int, int]:
        """Parse hex RGB string or named Azure color to (r, g, b)."""
        if color.lower() in AZURE_COLORS:
            return AZURE_COLORS[color.lower()]
        color = color.lstrip("#")
        if len(color) == 6:
            return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
        raise ValueError(f"Invalid color '{color}'. Use hex RGB or: {', '.join(AZURE_COLORS.keys())}")

    # ── Stencil discovery tools ──────────────────────────────

    def list_azure_services(self) -> list[str]:
        """Return all known Azure service keys from the stencil map."""
        return sorted(self._stencil_map.keys())

    def list_stencil_masters(self, stencil_name: str) -> list[str]:
        """List all master shape names in a given stencil file."""
        stencil = self._open_azure_stencil(stencil_name)
        masters = []
        for i in range(1, stencil.Masters.Count + 1):
            masters.append(stencil.Masters.Item(i).Name)
        return masters

    def open_stencil(self, stencil_name: str) -> str:
        """Open an Azure stencil by name. Returns the stencil document name."""
        stencil = self._open_azure_stencil(stencil_name)
        return stencil.Name

    # ── Document Management ──────────────────────────────────

    def create_diagram(self, template: str = "") -> str:
        """Create a new Visio diagram in landscape orientation (11×8.5). Returns the document name."""
        if template:
            doc = self.app.Documents.Add(template)
        else:
            doc = self.app.Documents.Add("")
        # Set landscape orientation per STYLE_GUIDE.md
        page = doc.Pages.Item(1)
        try:
            page.PageSheet.Cells("PageWidth").FormulaU = "11 in"
            page.PageSheet.Cells("PageHeight").FormulaU = "8.5 in"
            page.PageSheet.Cells("PrintPageOrientation").FormulaU = "2"  # landscape
        except Exception:
            pass
        return doc.Name

    def save_diagram(self, file_path: str) -> str:
        """Save the active diagram to a file path."""
        doc = self.active_doc
        if not doc:
            raise RuntimeError("No active document to save.")
        abs_path = os.path.abspath(file_path)
        doc.SaveAs(abs_path)
        return abs_path

    def close_diagram(self) -> str:
        """Close the active diagram."""
        doc = self.active_doc
        if not doc:
            return "No active document."
        name = doc.Name
        doc.Close()
        return f"Closed '{name}'."

    def list_open_diagrams(self) -> list[dict]:
        """List all open Visio documents."""
        results = []
        for i in range(1, self.app.Documents.Count + 1):
            doc = self.app.Documents.Item(i)
            results.append({
                "index": i,
                "name": doc.Name,
                "path": doc.FullName,
                "pages": doc.Pages.Count,
            })
        return results

    # ── Shape Operations ─────────────────────────────────────

    def add_shape(
        self,
        shape_type: str,
        x: float,
        y: float,
        text: str = "",
        width: float = 0,
        height: float = 0,
        fill_color: str | None = None,
    ) -> dict:
        """
        Add a basic shape to the active page with style-guide styling.

        Args:
            shape_type: Type name (rectangle, ellipse, diamond, etc.)
            x, y: Position in inches from bottom-left
            text: Text label for the shape
            width, height: Optional size override in inches
            fill_color: Hex RGB or named Azure color (e.g. "azure_blue", "teal", "FF8C00")
        Returns:
            dict with shape id, name, and position
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page. Create a diagram first.")

        try:
            master = self._get_master(shape_type)
            shape = page.Drop(master, x, y)
        except ValueError:
            shape = page.DrawRectangle(x - 0.75, y - 0.375, x + 0.75, y + 0.375)

        if text:
            shape.Text = text
        if width > 0:
            shape.Cells("Width").ResultIU = width
        if height > 0:
            shape.Cells("Height").ResultIU = height

        self._apply_shape_style(shape, fill_color)

        return {
            "id": shape.ID,
            "name": shape.Name,
            "text": text,
            "x": x,
            "y": y,
        }

    def add_azure_shape(
        self,
        service: str,
        x: float,
        y: float,
        text: str = "",
        width: float = 0,
        height: float = 0,
        fill_color: str | None = None,
    ) -> dict:
        """
        Add an Azure service icon from the real Azure stencils.

        Args:
            service: Azure service name (e.g. "azure/front-door", "SQL Database",
                     "vm-scale-sets"). Supports exact keys and fuzzy matching.
            x, y: Position in inches from bottom-left
            text: Optional label (defaults to master name)
            width, height: Optional size override in inches
            fill_color: Hex RGB or named Azure color
        Returns:
            dict with shape id, name, text, position, and resolved service info
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page. Create a diagram first.")

        resolved = self._resolve_azure_service(service)
        if not resolved:
            raise ValueError(
                f"Unknown Azure service '{service}'. "
                f"Use list_azure_services() to see available services."
            )

        stencil_name, master_name = resolved
        stencil = self._open_azure_stencil(stencil_name)
        master = stencil.Masters.Item(master_name)

        # Retry Drop with stencil re-open on transient COM errors
        last_err = None
        for attempt in range(3):
            try:
                pythoncom.CoInitialize()
                shape = page.Drop(master, x, y)
                break
            except Exception as e:
                last_err = e
                # Evict stale stencil cache and re-acquire master
                filename = f"{stencil_name}.vssx"
                self._stencils.pop(filename, None)
                stencil = self._open_azure_stencil(stencil_name)
                master = stencil.Masters.Item(master_name)
        else:
            raise RuntimeError(
                f"Failed to drop '{master_name}' after 3 attempts: {last_err}"
            )

        if text:
            shape.Text = text
        if width > 0:
            shape.Cells("Width").ResultIU = width
        if height > 0:
            shape.Cells("Height").ResultIU = height

        self._apply_shape_style(shape, fill_color)

        return {
            "id": shape.ID,
            "name": shape.Name,
            "text": text or master_name,
            "x": x,
            "y": y,
            "azure_service": f"azure/{_fuzzy_key(service).split('/')[-1]}",
            "stencil": stencil_name,
            "master": master_name,
        }

    def remove_shape(self, shape_id: int) -> str:
        """Remove a shape by its ID."""
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            if shape.ID == shape_id:
                name = shape.Name
                shape.Delete()
                return f"Deleted shape '{name}' (ID={shape_id})."
        raise ValueError(f"Shape with ID {shape_id} not found on active page.")

    def modify_shape(
        self,
        shape_id: int,
        text: str | None = None,
        x: float | None = None,
        y: float | None = None,
        width: float | None = None,
        height: float | None = None,
        fill_color: str | None = None,
    ) -> dict:
        """
        Modify properties of an existing shape.

        Args:
            shape_id: The shape's ID
            text: New text label
            x, y: New position (inches)
            width, height: New size (inches)
            fill_color: Fill color as RGB hex string e.g. "FF0000" for red
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")

        shape = None
        for i in range(1, page.Shapes.Count + 1):
            s = page.Shapes.Item(i)
            if s.ID == shape_id:
                shape = s
                break
        if not shape:
            raise ValueError(f"Shape with ID {shape_id} not found.")

        if text is not None:
            shape.Text = text
        if x is not None:
            shape.Cells("PinX").ResultIU = x
        if y is not None:
            shape.Cells("PinY").ResultIU = y
        if width is not None:
            shape.Cells("Width").ResultIU = width
        if height is not None:
            shape.Cells("Height").ResultIU = height
        if fill_color is not None:
            r, g, b = self._parse_color(fill_color)
            shape.Cells("FillForegnd").FormulaU = f"RGB({r},{g},{b})"

        return {
            "id": shape.ID,
            "name": shape.Name,
            "text": shape.Text,
        }

    def list_shapes(self) -> list[dict]:
        """List all shapes on the active page."""
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")
        results = []
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            try:
                results.append({
                    "id": shape.ID,
                    "name": shape.Name,
                    "text": shape.Text,
                    "x": round(shape.Cells("PinX").ResultIU, 2),
                    "y": round(shape.Cells("PinY").ResultIU, 2),
                    "width": round(shape.Cells("Width").ResultIU, 2),
                    "height": round(shape.Cells("Height").ResultIU, 2),
                })
            except Exception:
                results.append({"id": shape.ID, "name": shape.Name})
        return results

    # ── Connections ───────────────────────────────────────────

    def connect_shapes(
        self,
        from_shape_id: int,
        to_shape_id: int,
        label: str = "",
        connector_style: str = "straight",
        dashed: bool = False,
        bidirectional: bool = False,
    ) -> dict:
        """
        Connect two shapes with a styled connector per STYLE_GUIDE.md.

        Args:
            from_shape_id: Source shape ID
            to_shape_id: Target shape ID
            label: Optional text on the connector
            connector_style: "straight", "curved", or "right_angle"
            dashed: True for dashed line (failover, replication, secondary paths)
            bidirectional: True for arrows on both ends (replication links)
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")

        from_shape = to_shape = None
        for i in range(1, page.Shapes.Count + 1):
            s = page.Shapes.Item(i)
            if s.ID == from_shape_id:
                from_shape = s
            if s.ID == to_shape_id:
                to_shape = s

        if not from_shape:
            raise ValueError(f"Source shape ID {from_shape_id} not found.")
        if not to_shape:
            raise ValueError(f"Target shape ID {to_shape_id} not found.")

        connector = page.Drop(self.app.ConnectorToolDataObject, 0, 0)

        # Glue begin to source, end to target
        connector.Cells("BeginX").GlueTo(from_shape.Cells("PinX"))
        connector.Cells("EndX").GlueTo(to_shape.Cells("PinX"))

        # Routing style
        style_map = {
            "straight": 1,
            "curved": 2,
            "right_angle": 16,
        }
        route_style = style_map.get(connector_style.lower(), 1)
        try:
            connector.Cells("ShapeRouteStyle").ResultIU = route_style
        except Exception:
            pass

        # ── Style guide: arrowheads ──
        try:
            connector.Cells("EndArrow").FormulaU = str(CONNECTOR_END_ARROW)
            connector.Cells("EndArrowSize").FormulaU = str(CONNECTOR_ARROW_SIZE)
        except Exception:
            pass

        if bidirectional:
            try:
                connector.Cells("BeginArrow").FormulaU = str(CONNECTOR_END_ARROW)
                connector.Cells("BeginArrowSize").FormulaU = str(CONNECTOR_ARROW_SIZE)
            except Exception:
                pass

        # ── Style guide: line weight + rounding ──
        try:
            connector.Cells("LineWeight").FormulaU = f"{CONNECTOR_LINE_WEIGHT} pt"
            connector.Cells("Rounding").FormulaU = f"{CONNECTOR_ROUNDING} in"
        except Exception:
            pass

        # ── Style guide: dashed lines for secondary/replication paths ──
        if dashed:
            try:
                connector.Cells("LinePattern").FormulaU = str(CONNECTOR_DASHED_PATTERN)
            except Exception:
                pass

        # ── Label with 7pt font ──
        if label:
            connector.Text = label
            try:
                connector.Cells("Char.Size").FormulaU = f"{CONNECTOR_LABEL_FONT} pt"
            except Exception:
                pass

        return {
            "connector_id": connector.ID,
            "from": from_shape_id,
            "to": to_shape_id,
            "label": label,
        }

    def remove_connection(self, connector_id: int) -> str:
        """Remove a connector by its ID."""
        return self.remove_shape(connector_id)

    # ── Architecture Helpers ─────────────────────────────────

    def add_container(
        self,
        x: float,
        y: float,
        width: float,
        height: float,
        label: str = "",
        fill_color: str = "E6F3FF",
        transparency: float | None = None,
        rounding: float | None = None,
    ) -> dict:
        """
        Add a container/boundary rectangle styled per STYLE_GUIDE.md.
        Dashed border, 60% transparent, color-matched label at top.

        Args:
            x, y: Center position (inches)
            width, height: Size (inches)
            label: Title text for the container
            fill_color: Fill color as hex RGB or named Azure color
            transparency: Fill transparency 0.0–1.0 (default 0.60)
            rounding: Corner rounding in inches (default 0 = sharp)
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")

        trans_pct = int((transparency if transparency is not None else CONTAINER_FILL_TRANS / 100) * 100)
        round_in = rounding if rounding is not None else 0

        left = x - width / 2
        bottom = y - height / 2
        right = x + width / 2
        top = y + height / 2

        shape = page.DrawRectangle(left, bottom, right, top)

        r, g, b = self._parse_color(fill_color)
        shape.Cells("FillForegnd").FormulaU = f"RGB({r},{g},{b})"
        shape.Cells("FillForegndTrans").FormulaU = f"{trans_pct}%"

        if round_in > 0:
            shape.Cells("Rounding").FormulaU = f"{round_in} in"

        # Dashed border with darkened color for visibility
        dr, dg, db = max(r // 2, 0), max(g // 2, 0), max(b // 2, 0)
        shape.Cells("LinePattern").FormulaU = str(CONTAINER_LINE_PATTERN)
        shape.Cells("LineWeight").FormulaU = f"{CONTAINER_LINE_WEIGHT} pt"
        shape.Cells("LineColor").FormulaU = f"RGB({dr},{dg},{db})"

        shape.SendToBack()

        if label:
            shape.Text = label
            try:
                # Darkened label color for readability
                lr, lg, lb = max(r // 3, 0), max(g // 3, 0), max(b // 3, 0)
                shape.Cells("TxtPinY").FormulaU = CONTAINER_LABEL_PIN_Y
                shape.Cells("Char.Size").FormulaU = f"{CONTAINER_LABEL_FONT} pt"
                shape.Cells("Char.Style").FormulaU = "1"  # bold
                shape.Cells("Char.Color").FormulaU = f"RGB({lr},{lg},{lb})"
                shape.Cells("TopMargin").FormulaU = "6 pt"
                shape.Cells("VerticalAlign").FormulaU = "0"  # top
            except Exception:
                pass

        return {
            "id": shape.ID,
            "name": shape.Name,
            "label": label,
            "bounds": {"left": left, "bottom": bottom, "right": right, "top": top},
        }

    def add_text_label(self, x: float, y: float, text: str, font_size: int = 10) -> dict:
        """Add a floating text label at the given position."""
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")

        shape = page.DrawRectangle(x - 1, y - 0.2, x + 1, y + 0.2)
        shape.Text = text
        shape.Cells("LinePattern").FormulaU = "0"  # no border
        shape.Cells("FillPattern").FormulaU = "0"  # no fill
        shape.Cells("Char.Size").FormulaU = f"{font_size} pt"

        return {"id": shape.ID, "text": text, "x": x, "y": y}

    def add_tier_band(
        self,
        y: float,
        height: float,
        label: str = "",
        fill_color: str = "E6F3FF",
        page_width: float = 11.0,
        transparency: float | None = None,
        rounding: float | None = None,
    ) -> dict:
        """
        Add a horizontal tier band spanning the page width.
        70% transparent, no border, bold 8pt label on left margin.

        Args:
            y: Center Y position (inches)
            height: Band height (inches)
            label: Tier label (e.g. "Web Tier", "Data Tier")
            fill_color: Hex RGB or named Azure color
            page_width: Page width in inches (default 11 for landscape)
            transparency: Fill transparency 0.0–1.0 (default 0.70)
            rounding: Corner rounding in inches (default 0 = sharp)
        """
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")

        trans_pct = int((transparency if transparency is not None else TIER_BAND_FILL_TRANS / 100) * 100)
        round_in = rounding if rounding is not None else 0

        left = 0
        bottom = y - height / 2
        right = page_width
        top = y + height / 2

        shape = page.DrawRectangle(left, bottom, right, top)

        r, g, b = self._parse_color(fill_color)
        shape.Cells("FillForegnd").FormulaU = f"RGB({r},{g},{b})"
        shape.Cells("FillForegndTrans").FormulaU = f"{trans_pct}%"
        shape.Cells("LinePattern").FormulaU = "0"  # no border

        if round_in > 0:
            shape.Cells("Rounding").FormulaU = f"{round_in} in"

        shape.SendToBack()

        if label:
            shape.Text = label
            try:
                # Darkened label color for readability
                lr, lg, lb = max(r // 3, 0), max(g // 3, 0), max(b // 3, 0)
                shape.Cells("TxtPinX").FormulaU = "Width*0.06"
                shape.Cells("TxtPinY").FormulaU = "Height*0.5"
                shape.Cells("Para.HorzAlign").FormulaU = "0"  # left
                shape.Cells("Char.Size").FormulaU = f"{TIER_BAND_LABEL_FONT} pt"
                shape.Cells("Char.Style").FormulaU = "1"  # bold
                shape.Cells("Char.Color").FormulaU = f"RGB({lr},{lg},{lb})"
            except Exception:
                pass

        return {
            "id": shape.ID,
            "label": label,
            "bounds": {"left": left, "bottom": bottom, "right": right, "top": top},
        }

    # ── Page Operations ──────────────────────────────────────

    def add_page(self, name: str = "") -> dict:
        """Add a new page to the active document."""
        doc = self.active_doc
        if not doc:
            raise RuntimeError("No active document.")
        page = doc.Pages.Add()
        if name:
            page.Name = name
        return {"page_index": page.Index, "name": page.Name}

    def set_active_page(self, page_index: int) -> str:
        """Switch to a specific page by index (1-based)."""
        doc = self.active_doc
        if not doc:
            raise RuntimeError("No active document.")
        page = doc.Pages.Item(page_index)
        self.app.ActiveWindow.Page = page
        return f"Active page set to '{page.Name}' (index {page_index})."

    def list_pages(self) -> list[dict]:
        """List all pages in the active document."""
        doc = self.active_doc
        if not doc:
            raise RuntimeError("No active document.")
        return [
            {"index": i, "name": doc.Pages.Item(i).Name}
            for i in range(1, doc.Pages.Count + 1)
        ]

    # ── Export ───────────────────────────────────────────────

    def export_page(self, file_path: str) -> str:
        """Export the active page as PNG, SVG, or other image format."""
        page = self.active_page
        if not page:
            raise RuntimeError("No active page.")
        abs_path = os.path.abspath(file_path)
        page.Export(abs_path)
        return abs_path

    # ── Cleanup ──────────────────────────────────────────────

    def quit(self):
        """Release COM references (does NOT close Visio)."""
        self._stencils.clear()
        self._app = None
