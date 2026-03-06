# Visio MCP Server

An MCP (Model Context Protocol) server that exposes Microsoft Visio diagram operations as tools, designed for use with GitHub Copilot CLI and VS Code Agent Mode. Build Azure architecture diagrams through natural language.

## Features

- **Azure service icons** — 206 Azure services from official Visio stencils (auto-discovered)
- **Architecture helpers** — tier bands, containers, connectors with style-guide compliance
- **Shape operations** — add, modify, remove, connect, and list shapes
- **Multi-page support** — add pages, switch between them
- **Export** — PNG, SVG, JPG output

## Prerequisites

- **Microsoft Visio Professional** (installed and licensed)
- **Python 3.11+**
- **Azure Visio stencils** — download from [Microsoft Azure Architecture Icons](https://learn.microsoft.com/en-us/azure/architecture/icons/) and extract to your `My Shapes` folder

## Installation

```bash
cd visio-mcp
pip install -r requirements.txt
```

## Configuration

Add to your Copilot CLI MCP config (`~/.copilot/mcp-config.json`):

```json
{
  "mcpServers": {
    "visio": {
      "type": "stdio",
      "command": "python",
      "args": ["C:\\path\\to\\visio-mcp\\server.py"],
      "env": {}
    }
  }
}
```

### Whitelisting All Visio Tools

By default, Copilot CLI prompts for approval on each MCP tool call. To auto-approve all Visio MCP tools while keeping other tools at normal approval prompts, launch with:

```bash
copilot --allow-tool "visio"
```

To whitelist a specific tool only:

```bash
copilot --allow-tool "visio(add_azure_shape)"
```

To persist across sessions, add to `~/.copilot/config.json`:

```json
{
  "allowedTools": ["visio"]
}
```

## Available Tools

| Tool | Description |
|---|---|
| `create_diagram` | Create a new Visio diagram (landscape, 11×8.5 in) |
| `save_diagram` | Save to `.vsdx` file |
| `close_diagram` | Close without saving |
| `list_open_diagrams` | List all open documents |
| `add_shape` | Add basic shapes (rectangle, ellipse, diamond, etc.) |
| `add_azure_shape` | Add Azure service icons from official stencils |
| `remove_shape` | Remove a shape by ID |
| `modify_shape` | Change text, position, size, or color |
| `list_shapes` | List all shapes on the active page |
| `connect_shapes` | Connect two shapes with styled connectors |
| `remove_connection` | Remove a connector |
| `add_container` | Add a grouping boundary rectangle |
| `add_tier_band` | Add a full-width horizontal tier band |
| `add_text_label` | Add a floating text label |
| `list_azure_services` | List all 206 available Azure service keys |
| `list_stencil_masters` | List masters in a specific stencil |
| `open_stencil` | Open an Azure stencil by name |
| `add_page` | Add a new page |
| `set_active_page` | Switch to a page by index |
| `list_pages` | List all pages |
| `export_page` | Export page as image (PNG, SVG, JPG) |

## Azure Stencil Discovery

The server automatically discovers Azure stencils by querying Visio's COM properties:

1. `Visio.Application.MyShapesPath` — user stencils folder
2. `Visio.Application.GetBuiltInStencilFile()` — built-in content directory
3. `Visio.Application.Path` — install location
4. Windows Shell `CSIDL_PERSONAL` — real Documents folder (handles OneDrive redirect)

Discovered stencil directories are registered in Visio's `StencilPaths` for native COM resolution. No hardcoded paths.

## Style Guide

All shapes and connectors are automatically styled per `STYLE_GUIDE.md`:

- Shapes: rounded corners (0.06 in), 15% transparent fills
- Connectors: filled triangle arrowheads, 1 pt weight, 7 pt label font
- Containers: dashed border, 60% transparent, 9 pt label
- Tier bands: 70% transparent, bold 8 pt label

## Example

```
Create a 3-tier Azure architecture with Front Door, VM Scale Sets in 2 availability zones, and Azure SQL with replication
```

The server will create a professional Visio diagram with proper Azure icons, tier bands, containers, and styled connectors.

## License

MIT
