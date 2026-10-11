# illustrator-mcp

A small [Model Context Protocol](https://modelcontextprotocol.io) server that lets Claude (or any MCP client) drive a **locally installed Adobe Illustrator on Windows**.

It talks to Illustrator through COM and ExtendScript, entirely on your machine. Illustrator itself needs **no network connection** and no cloud API.

> Not affiliated with or endorsed by Adobe. "Adobe" and "Illustrator" are trademarks of Adobe Inc.

## Why

When an AI writes an `.ai`/SVG file from scratch, you usually get flattened shapes, broken text and no layers. With this server, **Illustrator does the drawing**, so you get:

- named layers, live editable text with your installed fonts, global swatches
- real `.ai` files saved by Illustrator itself
- a PNG preview loop, so the model can see what it drew and fix it
- the ability to open and edit your existing documents

## Requirements

- Windows with Adobe Illustrator installed (tested with Illustrator 2022 / v26.5; other versions that expose the `Illustrator.Application` COM object should work)
- Python 3.10+

## Install

```powershell
git clone https://github.com/kahrcitect/illustrator-mcp.git
cd illustrator-mcp
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

### Claude Code

```powershell
claude mcp add illustrator --scope user -- "C:\path\to\illustrator-mcp\.venv\Scripts\python.exe" "C:\path\to\illustrator-mcp\illustrator_mcp.py"
```

Restart Claude Code, then check with `claude mcp get illustrator`.

### Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "illustrator": {
      "command": "C:\\path\\to\\illustrator-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\illustrator-mcp\\illustrator_mcp.py"]
    }
  }
}
```

### Several Illustrator versions installed?

Pin one with the `ILLUSTRATOR_PROGID` environment variable, e.g. `Illustrator.Application.26` for 2022:

```powershell
claude mcp add illustrator --scope user -e ILLUSTRATOR_PROGID=Illustrator.Application.26 -- "...python.exe" "...illustrator_mcp.py"
```

## Tools

| Tool | What it does |
|---|---|
| `run_jsx` | Runs ExtendScript (ES3) code in Illustrator and returns the last value (objects as JSON). Errors come back as `ERROR: ...` instead of popping a dialog. |
| `get_document_info` | Open documents; for the active one: artboards, layers, swatches, selection, item counts. |
| `new_document` | New document in millimetres, RGB or CMYK. Artboard top-left is normalised to (0, 0). |
| `export_preview` | Renders an artboard to PNG and returns it as an image for visual checking. |
| `save_document` | Saves as `.ai`, editable `.pdf`, or exports `.svg`. |
| `open_document` | Opens an `.ai` / `.pdf` / `.svg` / `.eps` file. |

## Example prompts

- "Make an A3 landscape poster: dark background, my render centred, title top-left in Panchang."
- "In the open document, rename all layers to English and change every use of the red swatch to navy."
- "Export every artboard of this file as PNG and PDF into `C:\exports`."

## Notes and limits

- ExtendScript is **ES3**: no `JSON`, `let`/`const` or arrow functions. The server injects a small `__toJSON` helper.
- Coordinates are in points (1 mm = 2.834645 pt) and Illustrator's y axis points up.
- Some features are awkward or impossible from scripting (many live effects, some brushes, envelope distort). Workarounds are `app.executeMenuCommand(...)` or recorded Actions.
- If a modal dialog is open in Illustrator, scripts wait until you close it.

## Security

`run_jsx` executes **arbitrary ExtendScript** inside Illustrator. ExtendScript can read and write files on disk. Only connect this server to clients and prompts you trust, and review what the model runs if you are unsure.

## Türkçe

Windows'ta yerel olarak kurulu Adobe Illustrator'ı Claude ile kontrol etmeni sağlayan bir MCP sunucusu. Illustrator'ın internete bağlanması gerekmez, her şey COM + ExtendScript ile bilgisayarında çalışır. Kurulum için yukarıdaki **Install** adımlarını izle.

## License

MIT
