"""Illustrator MCP server (Windows, offline).

Talks to a locally installed Adobe Illustrator through COM and ExtendScript.
No network access is needed by Illustrator.
"""

import json
import os
import tempfile
import time

import pythoncom
import win32com.client
from mcp.server.mcpserver import Image, MCPServer

PREVIEW_DIR = os.path.join(tempfile.gettempdir(), "illustrator-mcp")
# Pin a version when several are installed, e.g. "Illustrator.Application.26" (2022).
PROGID = os.environ.get("ILLUSTRATOR_PROGID", "Illustrator.Application")
os.makedirs(PREVIEW_DIR, exist_ok=True)

mcp = MCPServer(
    "illustrator",
    instructions=(
        "Controls a local Adobe Illustrator via ExtendScript (ES3 JavaScript: "
        "no JSON object, no let/const, no arrow functions). Build editable work: named "
        "layers, live text, global swatches. After drawing, call export_preview to check "
        "the result visually. Units are points (1 mm = 2.834645 pt). Illustrator's y axis "
        "points up. Documents made by new_document have artboard top-left at (0, 0) with "
        "content at negative y; for other documents read artboardRect "
        "[left, top, right, bottom] from get_document_info and position relative to it."
    ),
)

# ES3 has no JSON; this helper is injected so scripts can return structured data.
JSX_HELPERS = r"""
function __toJSON(v) {
    if (v === null || v === undefined) return "null";
    var t = typeof v;
    if (t === "number") return isFinite(v) ? String(v) : "null";
    if (t === "boolean") return String(v);
    if (t === "string") {
        return '"' + v.replace(/\\/g, "\\\\").replace(/"/g, '\\"')
            .replace(/\n/g, "\\n").replace(/\r/g, "\\r").replace(/\t/g, "\\t") + '"';
    }
    if (v instanceof Array) {
        var a = [];
        for (var i = 0; i < v.length; i++) a.push(__toJSON(v[i]));
        return "[" + a.join(",") + "]";
    }
    var p = [];
    for (var k in v) if (v.hasOwnProperty(k)) p.push(__toJSON(String(k)) + ":" + __toJSON(v[k]));
    return "{" + p.join(",") + "}";
}
"""


def _app():
    pythoncom.CoInitialize()
    return win32com.client.Dispatch(PROGID)


def _run(code: str) -> str:
    """Run ExtendScript; errors come back as text instead of COM exceptions."""
    wrapped = (
        JSX_HELPERS
        + "var __prevUI = app.userInteractionLevel;"
        + "app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;"
        + "var __out;"
        + "try { var __r = eval(" + json.dumps(code) + ");"
        + " __out = (__r === undefined) ? 'OK' : ((typeof __r === 'object') ? __toJSON(__r) : String(__r)); }"
        + " catch (e) { __out = 'ERROR: ' + e.message + (e.line ? ' (line ' + e.line + ')' : ''); }"
        + " finally { app.userInteractionLevel = __prevUI; }"
        + "__out;"
    )
    result = _app().DoJavaScript(wrapped)
    return "" if result is None else str(result)


def _jsx_path(path: str) -> str:
    return json.dumps(os.path.abspath(path).replace("\\", "/"))


@mcp.tool()
def run_jsx(code: str) -> str:
    """Run ExtendScript (ES3) code in Illustrator and return the value of the last
    expression. Objects/arrays are returned as JSON. Errors are returned as
    'ERROR: ...'. Use `app.activeDocument` for the current document."""
    return _run(code)


@mcp.tool()
def get_document_info() -> str:
    """Describe the open documents and the active one: size, color mode, artboards,
    layers, swatches, selection and item counts."""
    return _run(r"""
(function () {
    var docs = [];
    for (var i = 0; i < app.documents.length; i++) docs.push(app.documents[i].name);
    if (app.documents.length === 0) return { openDocuments: docs, active: null };
    var d = app.activeDocument;
    var abs = [];
    for (var i = 0; i < d.artboards.length; i++) {
        var r = d.artboards[i].artboardRect;
        abs.push({ index: i, name: d.artboards[i].name, rect: r,
                   width: r[2] - r[0], height: r[1] - r[3] });
    }
    function layerInfo(l) {
        var subs = [];
        for (var j = 0; j < l.layers.length; j++) subs.push(layerInfo(l.layers[j]));
        return { name: l.name, visible: l.visible, locked: l.locked,
                 items: l.pageItems.length, sublayers: subs };
    }
    var layers = [];
    for (var i = 0; i < d.layers.length; i++) layers.push(layerInfo(d.layers[i]));
    var sel = [];
    var s = d.selection;
    if (s && s.length) for (var i = 0; i < s.length && i < 50; i++)
        sel.push({ type: s[i].typename, name: s[i].name, bounds: s[i].visibleBounds });
    var sw = [];
    for (var i = 0; i < d.swatches.length && i < 100; i++) sw.push(d.swatches[i].name);
    return {
        openDocuments: docs,
        active: {
            name: d.name, path: d.saved ? d.fullName.fsName : null, saved: d.saved,
            colorMode: String(d.documentColorSpace), rulerUnits: String(d.rulerUnits),
            artboards: abs, activeArtboard: d.artboards.getActiveArtboardIndex(),
            layers: layers, swatches: sw,
            counts: { pageItems: d.pageItems.length, textFrames: d.textFrames.length,
                      pathItems: d.pathItems.length, groups: d.groupItems.length,
                      placed: d.placedItems.length, raster: d.rasterItems.length },
            selection: sel
        }
    };
})()
""")


@mcp.tool()
def new_document(
    width_mm: float = 210,
    height_mm: float = 297,
    color_mode: str = "RGB",
    name: str = "",
) -> str:
    """Create a new document. color_mode: 'RGB' or 'CMYK'. Sizes in millimetres."""
    mode = "DocumentColorSpace.CMYK" if color_mode.upper() == "CMYK" else "DocumentColorSpace.RGB"
    pt = 2.834645669
    return _run(f"""
(function () {{
    var d = app.documents.add({mode}, {width_mm * pt}, {height_mm * pt});
    // Normalise so the artboard's top-left is (0, 0) and content goes to negative y.
    d.artboards[0].artboardRect = [0, 0, {width_mm * pt}, {-height_mm * pt}];
    {f"d.artboards[0].name = {json.dumps(name)};" if name else ""}
    return {{ name: d.name, width: {width_mm * pt}, height: {height_mm * pt} }};
}})()
""")


@mcp.tool()
def export_preview(artboard_index: int = -1, max_size_px: int = 1200) -> Image | str:
    """Render an artboard of the active document to PNG and return it as an image so
    the result can be checked visually. artboard_index -1 = active artboard."""
    out = os.path.join(PREVIEW_DIR, f"preview_{int(time.time() * 1000)}.png")
    res = _run(f"""
(function () {{
    var d = app.activeDocument;
    var idx = {artboard_index};
    if (idx < 0) idx = d.artboards.getActiveArtboardIndex();
    var prev = d.artboards.getActiveArtboardIndex();
    d.artboards.setActiveArtboardIndex(idx);
    var r = d.artboards[idx].artboardRect;
    var longest = Math.max(r[2] - r[0], r[1] - r[3]);
    var scale = Math.min(400, Math.max(5, {max_size_px} / longest * 100));
    var o = new ExportOptionsPNG24();
    o.artBoardClipping = true;
    o.antiAliasing = true;
    o.transparency = false;
    o.horizontalScale = scale;
    o.verticalScale = scale;
    d.exportFile(new File({_jsx_path(out)}), ExportType.PNG24, o);
    d.artboards.setActiveArtboardIndex(prev);
    return "OK";
}})()
""")
    if res.startswith("ERROR"):
        return res
    if not os.path.exists(out):
        return f"Export reported OK but no file at {out}"
    return Image(path=out)


@mcp.tool()
def save_document(path: str, format: str = "ai") -> str:
    """Save the active document. format: 'ai' (editable, keeps the document open
    under the new path), 'pdf' (Illustrator-editable PDF), or 'svg' (export copy)."""
    fmt = format.lower()
    if fmt == "ai":
        body = "var o = new IllustratorSaveOptions(); o.pdfCompatible = true; d.saveAs(f, o);"
    elif fmt == "pdf":
        body = "var o = new PDFSaveOptions(); o.preserveEditability = true; d.saveAs(f, o);"
    elif fmt == "svg":
        body = ("var o = new ExportOptionsSVG(); o.embedRasterImages = true; "
                "o.fontType = SVGFontType.OUTLINEFONT; d.exportFile(f, ExportType.SVG, o);")
    else:
        return "ERROR: format must be 'ai', 'pdf' or 'svg'"
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    return _run(f"""
(function () {{
    var d = app.activeDocument;
    var f = new File({_jsx_path(path)});
    {body}
    return f.fsName;
}})()
""")


@mcp.tool()
def open_document(path: str) -> str:
    """Open an .ai/.pdf/.svg/.eps file in Illustrator and make it active."""
    if not os.path.exists(path):
        return f"ERROR: file not found: {path}"
    return _run(f"app.open(new File({_jsx_path(path)})).name")


if __name__ == "__main__":
    mcp.run()
