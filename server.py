"""
Visio MCP Server
Exposes Microsoft Visio diagram operations as MCP tools via stdio transport.
Designed for use with GitHub Copilot CLI / VS Code Agent Mode.
All tools follow the conventions in STYLE_GUIDE.md automatically.
"""

import json
from mcp.server.fastmcp import FastMCP
from visio_client import VisioClient

mcp = FastMCP(name="Visio Diagram Server")
visio = VisioClient()


# ── Document Management ──────────────────────────────────────


@mcp.tool()
def create_diagram(template: str = "") -> str:
    """
    Create a new Visio diagram.
    The page is automatically set to **landscape** (11 × 8.5 in) per the style guide.

    Args:
        template: Optional Visio template name or path (e.g. "Basic Diagram.vstx").
                  Leave empty for a blank drawing.
    Returns:
        The name of the created document.
    """
    name = visio.create_diagram(template)
    return f"Created diagram: {name}"


@mcp.tool()
def save_diagram(file_path: str) -> str:
    """
    Save the active Visio diagram to a file.

    Args:
        file_path: Full path to save (e.g. "C:/Users/me/diagrams/arch.vsdx")
    """
    path = visio.save_diagram(file_path)
    return f"Saved to: {path}"


@mcp.tool()
def close_diagram() -> str:
    """Close the active Visio diagram without saving."""
    return visio.close_diagram()


@mcp.tool()
def list_open_diagrams() -> str:
    """List all open Visio documents with their page counts."""
    docs = visio.list_open_diagrams()
    if not docs:
        return "No open documents."
    return json.dumps(docs, indent=2)


# ── Shape Operations ─────────────────────────────────────────


@mcp.tool()
def add_shape(
    shape_type: str,
    x: float,
    y: float,
    text: str = "",
    width: float = 0,
    height: float = 0,
    fill_color: str | None = None,
) -> str:
    """
    Add a basic shape to the active Visio page.
    Shapes are automatically styled: rounded corners (0.06 in), semi-transparent fills (15%).

    For **Azure service icons** use `add_azure_shape` instead — it drops real stencil masters.

    Args:
        shape_type: Type of shape. Options: rectangle, square, ellipse, circle,
                    diamond, triangle, rounded_rectangle, star.
                    Or any master name from an open stencil.
        x: Horizontal position in inches from left edge.
        y: Vertical position in inches from bottom edge.
        text: Label text to display inside the shape.
        width: Width in inches (0 = default).
        height: Height in inches (0 = default).
        fill_color: Fill color as hex RGB (e.g. "FF0000") or named Azure color:
                    azure_blue, dark_blue, teal, orange, purple, green, red.

    Returns:
        JSON with shape id, name, text, and position.
    """
    result = visio.add_shape(shape_type, x, y, text, width, height, fill_color)
    return json.dumps(result, indent=2)


@mcp.tool()
def add_azure_shape(
    service: str,
    x: float,
    y: float,
    text: str = "",
    width: float = 0,
    height: float = 0,
    fill_color: str | None = None,
) -> str:
    """
    Add an **Azure service icon** from the official Azure Visio stencils.
    ALWAYS prefer this over add_shape for Azure architecture diagrams.

    The tool automatically opens the correct stencil file and drops the real
    Azure icon master. Shapes get rounded corners and semi-transparent fills.

    Args:
        service: Azure service name. Supports exact keys like "azure/front-door"
                 or fuzzy names like "Front Door", "SQL Database", "VM Scale Sets".
                 Use `list_azure_services` to see all 206 available services.
        x: Horizontal position in inches from left edge.
        y: Vertical position in inches from bottom edge.
        text: Optional label (defaults to the master shape name).
        width: Width in inches (0 = default stencil size).
        height: Height in inches (0 = default stencil size).
        fill_color: Optional fill color override (hex RGB or named Azure color).

    Returns:
        JSON with shape id, name, text, position, resolved service, stencil, and master.
    """
    result = visio.add_azure_shape(service, x, y, text, width, height, fill_color)
    return json.dumps(result, indent=2)


@mcp.tool()
def remove_shape(shape_id: int) -> str:
    """
    Remove a shape from the active page by its ID.

    Args:
        shape_id: The numeric ID of the shape (from add_shape or list_shapes).
    """
    return visio.remove_shape(shape_id)


@mcp.tool()
def modify_shape(
    shape_id: int,
    text: str | None = None,
    x: float | None = None,
    y: float | None = None,
    width: float | None = None,
    height: float | None = None,
    fill_color: str | None = None,
) -> str:
    """
    Modify properties of an existing shape.

    Args:
        shape_id: The shape's numeric ID.
        text: New text label (or null to keep current).
        x: New X position in inches (or null to keep).
        y: New Y position in inches (or null to keep).
        width: New width in inches (or null to keep).
        height: New height in inches (or null to keep).
        fill_color: Fill color as hex RGB (e.g. "FF0000") or named Azure color:
                    azure_blue, dark_blue, teal, orange, purple, green, red.
    """
    result = visio.modify_shape(shape_id, text, x, y, width, height, fill_color)
    return json.dumps(result, indent=2)


@mcp.tool()
def list_shapes() -> str:
    """
    List all shapes on the active Visio page.
    Returns JSON array with each shape's id, name, text, position, and size.
    """
    shapes = visio.list_shapes()
    if not shapes:
        return "No shapes on the active page."
    return json.dumps(shapes, indent=2)


# ── Connections ──────────────────────────────────────────────


@mcp.tool()
def connect_shapes(
    from_shape_id: int,
    to_shape_id: int,
    label: str = "",
    connector_style: str = "straight",
    dashed: bool = False,
    bidirectional: bool = False,
) -> str:
    """
    Connect two shapes with a styled connector line.
    Automatically applies style-guide rules: filled triangle arrowheads,
    1 pt line weight, 0.15 in rounding, 7 pt label font.

    Args:
        from_shape_id: ID of the source shape.
        to_shape_id: ID of the target shape.
        label: Optional text label on the connector (e.g. "HTTPS", "TDS").
        connector_style: One of "straight", "curved", or "right_angle".
        dashed: If True, uses dashed line (for failover, replication, secondary paths).
        bidirectional: If True, arrows on both ends (for replication links).
    """
    result = visio.connect_shapes(
        from_shape_id, to_shape_id, label, connector_style, dashed, bidirectional
    )
    return json.dumps(result, indent=2)


@mcp.tool()
def remove_connection(connector_id: int) -> str:
    """
    Remove a connector by its ID.

    Args:
        connector_id: The numeric ID of the connector shape.
    """
    return visio.remove_connection(connector_id)


# ── Architecture Helpers ─────────────────────────────────────


@mcp.tool()
def add_container(
    x: float,
    y: float,
    width: float,
    height: float,
    label: str = "",
    fill_color: str = "E6F3FF",
    transparency: float | None = None,
    rounding: float | None = None,
) -> str:
    """
    Add a container/boundary rectangle for visually grouping shapes.
    Styled per guide: dashed border, 60% transparent, 9 pt color-matched label at top.

    Args:
        x: Center X position in inches.
        y: Center Y position in inches.
        width: Width in inches.
        height: Height in inches.
        label: Title text displayed at the top of the container.
        fill_color: Fill color as hex RGB or named Azure color.
                    Defaults to "E6F3FF" (light blue).
        transparency: Fill transparency from 0.0 (opaque) to 1.0 (invisible).
                      Defaults to 0.60 if not specified.
        rounding: Corner rounding in inches (e.g. 0.12). Defaults to 0 (sharp corners).
    """
    result = visio.add_container(x, y, width, height, label, fill_color, transparency, rounding)
    return json.dumps(result, indent=2)


@mcp.tool()
def add_tier_band(
    y: float,
    height: float,
    label: str = "",
    fill_color: str = "E6F3FF",
    transparency: float | None = None,
    rounding: float | None = None,
) -> str:
    """
    Add a horizontal tier band spanning the full page width.
    Styled per guide: 70% transparent, no border, bold 8 pt label on left margin.
    Use for separating architecture tiers (e.g. "Web Tier", "App Tier", "Data Tier").

    Args:
        y: Center Y position in inches.
        height: Band height in inches.
        label: Tier label text (e.g. "Ingress", "Compute", "Data").
        fill_color: Fill color as hex RGB or named Azure color.
        transparency: Fill transparency from 0.0 (opaque) to 1.0 (invisible).
                      Defaults to 0.70 if not specified.
        rounding: Corner rounding in inches (e.g. 0.12). Defaults to 0 (sharp corners).
    """
    result = visio.add_tier_band(y, height, label, fill_color, transparency=transparency, rounding=rounding)
    return json.dumps(result, indent=2)


@mcp.tool()
def add_text_label(x: float, y: float, text: str, font_size: int = 10) -> str:
    """
    Add a floating text label (no border, no fill) at the given position.

    Args:
        x: X position in inches.
        y: Y position in inches.
        text: The text to display.
        font_size: Font size in points (default 10).
    """
    result = visio.add_text_label(x, y, text, font_size)
    return json.dumps(result, indent=2)


# ── Stencil Discovery ───────────────────────────────────────


@mcp.tool()
def list_azure_services() -> str:
    """
    List all 206 available Azure service keys that can be used with add_azure_shape.
    Returns a sorted JSON array of service identifiers like "azure/front-door",
    "azure/sql-database", "azure/vm-scale-sets", etc.
    """
    services = visio.list_azure_services()
    return json.dumps(services, indent=2)


@mcp.tool()
def list_stencil_masters(stencil_name: str) -> str:
    """
    List all master shape names in a given Azure stencil.
    Use this to discover what icons are available in a specific stencil.

    Args:
        stencil_name: Stencil logical name (e.g. "Azure-Databases", "Azure-Compute",
                      "Azure-Networking", "Azure-Web").
    """
    masters = visio.list_stencil_masters(stencil_name)
    return json.dumps(masters, indent=2)


@mcp.tool()
def open_stencil(stencil_name: str) -> str:
    """
    Open an Azure stencil by name so its masters become available.

    Args:
        stencil_name: Stencil logical name (e.g. "Azure-Databases").
    """
    name = visio.open_stencil(stencil_name)
    return f"Opened stencil: {name}"


# ── Page Operations ──────────────────────────────────────────


@mcp.tool()
def add_page(name: str = "") -> str:
    """
    Add a new page to the active Visio document.

    Args:
        name: Optional name for the new page (e.g. "Network Layer").
    """
    result = visio.add_page(name)
    return json.dumps(result, indent=2)


@mcp.tool()
def set_active_page(page_index: int) -> str:
    """
    Switch to a specific page by its index (1-based).

    Args:
        page_index: The page number to activate (1 = first page).
    """
    return visio.set_active_page(page_index)


@mcp.tool()
def list_pages() -> str:
    """List all pages in the active Visio document."""
    pages = visio.list_pages()
    return json.dumps(pages, indent=2)


# ── Export ───────────────────────────────────────────────────


@mcp.tool()
def export_page(file_path: str) -> str:
    """
    Export the active page as an image (PNG, SVG, JPG, etc.).
    The format is determined by the file extension.

    Args:
        file_path: Output path, e.g. "C:/Users/me/diagram.png" or "output.svg".
    """
    path = visio.export_page(file_path)
    return f"Exported to: {path}"


# ── Entry Point ──────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run(transport="stdio")
