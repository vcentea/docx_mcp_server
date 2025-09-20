# DOCX MCP Server

A lightweight Python MCP server that provides high-level tools to read and modify Microsoft Word (`.docx`) documents using a structured JSON representation. It supports converting DOCX → JSON, deleting/adding/editing elements, and applying multiple operations atomically with automatic versioning.

- **Convert DOCX to JSON** with semantic text formatting references
- **Edit element content/properties** by ID
- **Add or delete elements** (single or multiple)
- **Combined operations** (add + edit + delete in one call)
- **Automatic versioning** (`document.v1.docx`, `document.v2.docx`, ...)
- **Flexible responses**: minimal, ID mapping, or full updated JSON

Repository: `vcentea/docx_mcp_server` (`main`)  
GitHub: `https://github.com/vcentea/docx_mcp_server`

## Requirements

- Python 3.10+
- System dependencies: none (pure Python)

Install Python dependencies:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

`requirements.txt` contains:
- `mcp` (FastMCP server runtime)
- `lxml` (WordprocessingML parsing)
- `python-docx` (DOCX reconstruction)
- `uvicorn`, `sse-starlette` (optional HTTP/SSE transports)
- `pydantic` (imported by server)

## Quick Start (STDIO)

Run the server over stdio (default):

```bat
.venv\Scripts\activate
python mcp_server.py
```

Point your MCP client (e.g., IDE/agent) to spawn the above command. The server name is `contract-docx-mcp`.

## Network Modes (Optional)

The server can run with different transports using environment variables:

- `MCP_TRANSPORT` = `stdio` (default) | `sse` | `http` | `ws`
- `MCP_HOST` (default `127.0.0.1`)
- `MCP_PORT` (default `8765`)
- `MCP_PATH` (e.g., `/mcp` for HTTP)
- `MCP_SSE_PATH` (default `/sse`) and `MCP_MESSAGE_PATH` (SSE messaging base)

Examples:

```bat
set MCP_TRANSPORT=sse
set MCP_PORT=8765
python mcp_server.py
```

```bat
set MCP_TRANSPORT=http
set MCP_PATH=/mcp
python mcp_server.py
```

If a network transport fails, the server falls back to stdio.

## Tools (API)

All tools are exposed via MCP. The JSON schema shown is conceptual and may be invoked according to your MCP client format.

### 1) get_document_as_json
Convert a `.docx` file into structured JSON with a `textProperties` registry and element IDs.

Args:
- `docx_path: string` (required)
- `output_json_path?: string`
- `return_json?: boolean` (default: true)

Returns (if `return_json=true`):
- `{ source_file, textProperties, body }`

Example:
```json
{
  "docx_path": "contract.docx",
  "return_json": true
}
```

### 2) get_document_text_properties
Extract only the text formatting registry.

Args:
- `docx_path: string`
- `return_json?: boolean` (default: true)

Returns: `{ source_file, textProperties, properties_count }`

### 3) delete_elements
Delete one or more elements by ID.

Args:
- `docx_path: string`
- `element_ids: string | string[]`
- `response_format?: "minimal" | "id_mapping" | "full_document"` (default: `minimal`)
- `output_docx_path?: string`

Returns: `{ output_docx_path, version, success, deleted_element_ids, deleted_count, id_mapping?, updated_document? }`

Example:
```json
{
  "docx_path": "contract.docx",
  "element_ids": ["p-3", "t-1"],
  "response_format": "id_mapping"
}
```

### 4) add_elements
Insert one or more new elements with optional formatting. Positions: `after`, `before`, `end`.

Args:
- `docx_path: string`
- `new_elements: string | object | (string|object)[]`
- `position: "after" | "before" | "end"`
- `text_properties_ref?: string`
- `reference_element_id?: string` (required for `after`/`before`)
- `response_format?: "minimal" | "id_mapping" | "full_document"` (default: `minimal`)
- `output_docx_path?: string`

Returns: `{ output_docx_path, version, success, added_position, reference_element_id?, added_count, id_mapping?, updated_document? }`

Examples:
- Append simple text using a known format name:
```json
{
  "docx_path": "contract.docx",
  "new_elements": "New paragraph text",
  "position": "end",
  "text_properties_ref": "default_text_format",
  "response_format": "id_mapping"
}
```
- Insert multiple paragraphs after an element:
```json
{
  "docx_path": "contract.docx",
  "new_elements": ["First", "Second"],
  "position": "after",
  "reference_element_id": "p-2"
}
```

### 5) edit_element_content
Edit content or properties of a single element by ID. For text content, you can optionally apply a `text_properties_ref`.

Args:
- `docx_path: string`
- `element_id: string`
- `property_path: string` (e.g., `content`, `pPr.jc`)
- `new_value: any`
- `text_properties_ref?: string`
- `response_format?: "minimal" | "id_mapping" | "full_document"`
- `output_docx_path?: string`

Returns: `{ output_docx_path, version, success, edited_element_id, property_path, applied_text_properties_ref?, id_mapping?, updated_document? }`

Example (replace text content and apply formatting):
```json
{
  "docx_path": "contract.docx",
  "element_id": "p-1",
  "property_path": "content",
  "new_value": "Updated Title",
  "text_properties_ref": "heading_1_bold_arial_12pt_format"
}
```

### 6) edit_document (combined operations)
Apply deletions → edits → additions atomically in one call.

Args:
- `docx_path: string`
- `edits?: EditOperation[]`
- `additions?: AdditionOperation[]`
- `deletions?: string[]`
- `response_format?: "minimal" | "id_mapping" | "full_document"`
- `output_docx_path?: string`

Returns: `{ output_docx_path, version, success, operations_applied: { deletions, edits, additions, total }, id_mapping?, updated_document? }`

Example:
```json
{
  "docx_path": "contract.docx",
  "deletions": ["p-3"],
  "edits": [{
    "element_id": "p-1",
    "property_path": "content",
    "new_value": "Updated",
    "text_properties_ref": "heading_1_bold_arial_12pt_format"
  }],
  "additions": [{
    "elements": "Conclusion",
    "position": "end",
    "text_properties_ref": "default_text_format"
  }],
  "response_format": "id_mapping"
}
```

## Versioning & ID Mapping
- Every write creates a new versioned file near the source: `contract.v1.docx`, `contract.v2.docx`, ...
- Choose `response_format: "id_mapping"` or `"full_document"` to receive a mapping of deleted/new IDs and (optionally) the updated JSON.

## Minimal File Set
Only these files are required to run the server:

- `mcp_server.py`
- `patch_json.py`
- `docx2json_runs_only.py`
- `docx2json_orchestrate.py`
- `json2docx_structured.py`
- `requirements.txt`
- `README.md`

Everything else can be excluded (old versions, environment files, legacy folders, tests, samples).

## Push to GitHub (Windows)
Prepare a clean folder with only the necessary files and push to `main`:

```bat
set SRC="E:\Google Drive AInnovate\vlad\_PROJECTS\contract_manager"
set DEST="E:\Google Drive AInnovate\vlad\_PROJECTS\docx_mcp_server_clean"

mkdir %DEST%
copy %SRC%\mcp_server.py %DEST%\
copy %SRC%\patch_json.py %DEST%\
copy %SRC%\docx2json_runs_only.py %DEST%\
copy %SRC%\docx2json_orchestrate.py %DEST%\
copy %SRC%\json2docx_structured.py %DEST%\
copy %SRC%\requirements.txt %DEST%\
copy %SRC%\README.md %DEST%\

cd /d %DEST%
git init
git branch -M main
git add .
git commit -m "feat: initial DOCX MCP server minimal set"
git remote add origin https://github.com/vcentea/docx_mcp_server.git
rem If the remote has existing commits (e.g., LICENSE), pull and merge first:
rem git pull origin main --allow-unrelated-histories --no-edit

git push -u origin main
```

If `git push` is rejected due to non-fast-forward, either pull with `--allow-unrelated-histories` as shown, or force push if appropriate for your repository policy (`git push -f`).

## License
This project is published under the MIT License. See the repository for details.
