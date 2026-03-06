# Visio Architecture Diagram — Style Guide

## Stencils

- Always use **real Azure stencil icons** from Visio's built-in stencils (e.g. `AZURECOMPUTE_M.VSSX`, `AZURENETWORKING_M.VSSX`, `AZUREDATABASES_M.VSSX`, `AZURESTORAGE_M.VSSX`, etc.)
- Never use generic rectangles when an official Azure shape exists
- Discover available masters with `stencil.Masters.Item(i).Name` before building

## Shapes

- **Rounded corners** on all component shapes: `Rounding = 0.06 in`
- **Semi-transparent fills**: `FillForegndTrans = 15%`
- Use the **Azure brand color palette**:

| Color           | RGB              | Use For                          |
|-----------------|------------------|----------------------------------|
| Azure Blue      | `0, 120, 215`    | Load Balancers, SQL, general     |
| Dark Blue       | `0, 78, 152`     | Internal LBs, Private Endpoints  |
| Teal            | `0, 178, 148`    | Front Door, CDN                  |
| Orange          | `255, 140, 0`    | Web tier, public-facing compute  |
| Purple          | `135, 100, 184`  | App tier, middleware              |
| Green           | `122, 184, 0`    | Storage, data lake               |
| Red             | `232, 17, 35`    | Alerts, errors, critical paths   |

## Connectors

- **Filled triangle arrowheads**: `EndArrow = 4`, `EndArrowSize = 2`
- **Slight rounding** at bends: `Rounding = 0.15 in`
- **Line weight**: `1 pt`
- **Dashed lines** for failover, replication, or secondary paths: `LinePattern = 2`
- **Bidirectional arrows** for replication links (set both `BeginArrow` and `EndArrow`)
- **Label font**: `7 pt`

## Containers / Zones

- **Dashed border**: `LinePattern = 2`, `LineWeight = 1 pt`
- **Fill**: color-matched, `60% transparent`
- **Label position**: top of container (`TxtPinY = Height*0.96`)
- **Label font**: `9 pt`, color-matched to border
- For tier bands: `70% transparent`, no border, with **bold 8pt labels** on the left margin

## Page Layout

- **Landscape orientation**: 11 × 8.5 inches
- **Content centered** with ~1 inch margins on all sides
- **Top-to-bottom flow**: users → frontend → web → app → data
- **Left/right symmetry** for availability zones or redundant paths
