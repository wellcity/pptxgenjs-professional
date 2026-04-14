# pptxgenjs-professional

Professional PowerPoint generation skill using pptxgenjs — featuring the **Midnight Executive** design system (Microsoft-inspired deep blue palette), unified Header/Footer, data cards, architecture diagrams, flowcharts, and a complete icon system.

## Quick Start

```bash
# Install dependencies (once)
cd /tmp && npm init -y && npm install pptxgenjs react react-dom sharp react-icons

# Copy template
cp references/template.js my-presentation.js

# Edit content, then generate
node my-presentation.js
# Output: output.pptx
```

## What's Included

| File | Description |
|------|-------------|
| `SKILL.md` | Complete design system documentation |
| `references/template.js` | Minimal starter template |
| `references/microsoft-agent-framework.js` | Full 9-slide example (Microsoft Agent Framework topic) |

## Design System

### Color Palette: Midnight Executive

| Name | Hex | Usage |
|------|-----|-------|
| navy | `1E2761` | Primary dark |
| teal | `0078D4` | Main accent (Microsoft blue) |
| mid | `2E4A8F` | Medium blue |
| ice | `CADCFC` | Light accent |
| coral | `E85D4C` | Secondary accent |
| gold | `F5A623` | Tertiary accent |
| success | `10B981` | Green |
| purple | `6B5CE7` | Purple |
| darkBg | `0D1B3E` | Dark background (title/conclusion pages) |

### Standard Slide Layout

Every slide includes:
- **Header** (navy, white text, 1.1" height)
- **Footer** (navy, ice text, 0.4" height)
- **Content area** (offWhite or white card backgrounds)

## Slide Types

| Type | Description |
|------|-------------|
| Title page | Dark background + decorative circles |
| Definition + 3 pillars | Full-width definition + 3 icon cards |
| Left text + right diagram | Feature list on left, architecture on right |
| 3×2 grid cards | 6 cards with icons, colors, stats |
| Step flow | Vertical timeline with numbered circles |
| Stats dashboard | 2×2 or 2×3 stat boxes |
| Summary | Dark background + key points list |

## QA Checklist

```bash
python3 -m markitdown output.pptx 2>/dev/null | grep -iE "xxxx|lorem|ipsum|placeholder" || echo "✅ No placeholders"
```

- [ ] All text replaced, no "XXXX" / "Lorem" placeholders
- [ ] Each slide has visual elements (not just plain text)
- [ ] Dark/light background alternation (title/conclusion = dark, content = light)
- [ ] Different layouts across pages (avoid repetition)

## Dependencies

```json
{
  "pptxgenjs": "^4.0.1",
  "react": "^19.2.0",
  "react-dom": "^19.2.0",
  "react-icons": "^5.6.0",
  "sharp": "^0.34.0"
}
```
