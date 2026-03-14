# LibreOffice UNO Mail Merge — Technical Research

## 1. Word MERGEFIELD Internals

### How MERGEFIELDs Are Stored in .docx XML

A `.docx` file is a ZIP archive containing XML files. The main document body lives in `word/document.xml`. Word merge fields are stored as **field codes** using the Office Open XML (OOXML) field markup.

There are two representations:

#### Simple Field (w:fldSimple)

The compact form — rarely used by modern Word for merge fields:

```xml
<w:fldSimple w:instr=" MERGEFIELD CustomerName \* MERGEFORMAT ">
  <w:r>
    <w:t>«CustomerName»</w:t>
  </w:r>
</w:fldSimple>
```

#### Complex Field (w:fldChar) — The Common Form

This is how Word actually stores merge fields in practice. The field is split across multiple `<w:r>` (run) elements:

```xml
<!-- 1. BEGIN marker -->
<w:r>
  <w:rPr>
    <w:sz w:val="22"/>
  </w:rPr>
  <w:fldChar w:fldCharType="begin"/>
</w:r>

<!-- 2. Field instruction -->
<w:r>
  <w:rPr>
    <w:sz w:val="22"/>
  </w:rPr>
  <w:instrText xml:space="preserve"> MERGEFIELD CustomerName \* MERGEFORMAT </w:instrText>
</w:r>

<!-- 3. SEPARATE marker (divides instruction from display value) -->
<w:r>
  <w:fldChar w:fldCharType="separate"/>
</w:r>

<!-- 4. Display value (what the user sees: «CustomerName») -->
<w:r w:rsidR="00FA6E12">
  <w:rPr>
    <w:noProof/>
  </w:rPr>
  <w:t>«CustomerName»</w:t>
</w:r>

<!-- 5. END marker -->
<w:r>
  <w:rPr>
    <w:noProof/>
  </w:rPr>
  <w:fldChar w:fldCharType="end"/>
</w:r>
```

**Key structure:**
- `w:fldChar[@fldCharType="begin"]` — opens the field
- `w:instrText` — contains the field instruction (e.g., `MERGEFIELD CustomerName \* MERGEFORMAT`)
- `w:fldChar[@fldCharType="separate"]` — separates instruction from display
- Display text (e.g., `«CustomerName»`) — what's shown when field codes are hidden
- `w:fldChar[@fldCharType="end"]` — closes the field

**Critical gotcha:** Word may split the `w:instrText` across multiple `<w:r>` elements, especially when formatting changes mid-field. For example:

```xml
<w:r><w:instrText xml:space="preserve"> MERGE</w:instrText></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:instrText>FIELD Custom</w:instrText></w:r>
<w:r><w:instrText>erName \* MERGEFORMAT </w:instrText></w:r>
```

This is why simple text-replacement approaches fail — you need to concatenate all `w:instrText` content between `begin` and `separate` to get the full field instruction.

### Field Instruction Syntax

```
MERGEFIELD FieldName [\* switches]
```

Common switches:
- `\* MERGEFORMAT` — preserve formatting from the template
- `\* Upper` / `\* Lower` / `\* FirstCap` — text case transforms
- `\b "text"` — text before (inserted only if field has data)
- `\f "text"` — text after

### Conditional and Nested Fields

Word supports conditional merge fields:

```xml
<!-- IF field with nested MERGEFIELD -->
<w:instrText> IF </w:instrText>
<!-- nested MERGEFIELD begin -->
<w:fldChar w:fldCharType="begin"/>
<w:instrText> MERGEFIELD Status </w:instrText>
<w:fldChar w:fldCharType="separate"/>
<w:instrText>«Status»</w:instrText>
<w:fldChar w:fldCharType="end"/>
<!-- continuation of IF -->
<w:instrText> = "Active" "Yes" "No" </w:instrText>
```

Other field types that appear in mail merge templates:
- `IF` — conditional text
- `NEXT` — advance to next record (for labels/multi-record pages)
- `NEXTIF` — conditional record advance
- `SET` / `ASK` — variable assignment
- `DATE`, `PAGE`, `NUMPAGES` — document metadata

### How LibreOffice Interprets These

When LibreOffice opens a `.docx` file:

1. **The OOXML import filter parses the XML** and converts field codes to LibreOffice's internal field representation
2. **MERGEFIELD codes become `com.sun.star.text.fieldmaster.Database` fields** — LibreOffice maps them to its own database field infrastructure
3. **The field names are preserved** but the data source association may be lost (since the `.docx` doesn't contain a LibreOffice-compatible data source reference)
4. **IF/conditional fields** have mixed support — simple `IF MERGEFIELD = value` works, but complex nesting may fail
5. **NEXT/NEXTIF fields** are generally supported for multi-record-per-page layouts

**Important:** LibreOffice understands the field *codes* but needs a registered data source to actually execute the merge. The template alone isn't enough.

---

## 2. LibreOffice UNO Mail Merge Service

### The `com.sun.star.text.MailMerge` Service

This is the primary API entry point for programmatic mail merge. It implements:
- `com.sun.star.task.XJob` — the `execute()` method that runs the merge
- `com.sun.star.util.XCancellable` — cancel a running merge
- `com.sun.star.beans.XPropertySet` — get/set merge properties
- `com.sun.star.text.XMailMergeBroadcaster` — event notifications
- `com.sun.star.sdb.DataAccessDescriptor` — data source configuration (inherited)

### Required Properties

| Property | Type | Description |
|----------|------|-------------|
| `DocumentURL` | string | `file:///path/to/template.docx` — the template document |
| `DataSourceName` | string | Name of registered data source |
| `Command` | string | Table/query name within the data source |
| `CommandType` | long | `0` = TABLE, `1` = QUERY, `2` = SQL COMMAND |
| `OutputType` | short | `1` = PRINTER, `2` = FILE, `3` = MAIL, `4` = SHELL |

### Output Properties (when OutputType = FILE)

| Property | Type | Description |
|----------|------|-------------|
| `OutputURL` | string | `file:///path/to/output/directory/` |
| `SaveFilter` | string | Output format filter name |
| `SaveAsSingleFile` | boolean | `True` = one file, `False` = file per record |
| `FileNameFromColumn` | boolean | Use a DB column for file names |
| `FileNamePrefix` | string | Column name for file names, or prefix if not from column |

### Additional Properties

| Property | Type | Description |
|----------|------|-------------|
| `Filter` | string | SQL WHERE clause to filter records |
| `EscapeProcessing` | boolean | Whether to parse SQL |
| `SinglePrintJobs` | boolean | Separate print jobs per document |
| `Model` | XModel | Alternative to DocumentURL — use an already-opened document |
| `ResultSet` | XResultSet | Alternative to DataSourceName — provide data directly |
| `ActiveConnection` | XConnection | Alternative — provide an existing DB connection |

### MailMergeType Constants

```
com.sun.star.text.MailMergeType.PRINTER  = 1  # Send to printer
com.sun.star.text.MailMergeType.FILE     = 2  # Save to files
com.sun.star.text.MailMergeType.MAIL     = 3  # Send as email
com.sun.star.text.MailMergeType.SHELL    = 4  # Return XTextDocument (since LO 4.4)
```

### SaveFilter Values

| Filter Name | Output Format |
|------------|---------------|
| `"writer8"` | ODF (.odt) |
| `"MS Word 2007 XML"` | OOXML (.docx) |
| `"writer_pdf_Export"` | PDF |
| `"HTML (StarWriter)"` | HTML |
| `"Rich Text Format"` | RTF |

### Executing the Merge

```python
# The execute() method takes a tuple of NamedValue (can be empty)
oMailMerge.execute(())  # Note: tuple, not list!
```

**Critical gotcha:** Passing a Python `list` (`[]`) instead of a `tuple` (`()`) causes a `RuntimeException` because UNO expects a sequence with a `getTypes()` method. Always use `((),)` or `(())` — the exact syntax that works is:

```python
oMailMerge.execute(())
```

---

## 3. Registering a CSV Data Source

This is the hardest part of programmatic mail merge. LibreOffice needs a **registered data source** — you can't just point it at a CSV file directly.

### Approach 1: Create and Register via DatabaseContext

```python
import uno
from com.sun.star.beans import PropertyValue

def register_csv_datasource(ctx, csv_dir, ds_name):
    """
    Register a directory of CSV files as a data source.
    Each CSV file becomes a "table" in the data source.
    """
    smgr = ctx.ServiceManager
    
    # Get the DatabaseContext (manages all registered data sources)
    dbContext = smgr.createInstanceWithContext(
        "com.sun.star.sdb.DatabaseContext", ctx
    )
    
    # Create a new data source
    ds = dbContext.createInstance()
    
    # Configure for flat file (CSV) access
    # The URL points to the DIRECTORY containing CSV files
    ds.URL = "sdbc:flat:" + uno.systemPathToFileUrl(csv_dir)
    
    # Set CSV-specific connection properties
    info = []
    
    # Field delimiter
    p1 = PropertyValue()
    p1.Name = "Extension"
    p1.Value = "csv"
    info.append(p1)
    
    p2 = PropertyValue()
    p2.Name = "HeaderLine"
    p2.Value = True
    info.append(p2)
    
    p3 = PropertyValue()
    p3.Name = "FieldDelimiter"
    p3.Value = ","
    info.append(p3)
    
    p4 = PropertyValue()
    p4.Name = "StringDelimiter"
    p4.Value = '"'
    info.append(p4)
    
    p5 = PropertyValue()
    p5.Name = "CharSet"
    p5.Value = "UTF-8"
    info.append(p5)
    
    ds.Info = tuple(info)
    
    # Register under a name
    dbContext.registerObject(ds_name, ds)
    
    # The data source must be stored (creates a .odb file in LO's user profile)
    ds.DatabaseDocument.store()
    
    return ds_name
```

### Approach 2: Create an .odb File Manually

An alternative is to create a Base `.odb` file (which is a ZIP containing XML) that points to a CSV directory. This avoids needing UNO for registration but adds file management complexity.

### Approach 3: Use ResultSet Directly (Bypass Data Source Registration)

Instead of registering a data source, provide data directly via the `ResultSet` property:

```python
# This approach avoids data source registration entirely
# but requires more complex code to create an XResultSet
oMailMerge.ActiveConnection = connection
oMailMerge.ResultSet = resultSet
```

This is more complex but avoids the data source registration dance. It requires creating an in-memory result set, which itself is non-trivial in UNO.

### Recommended Approach for Lola

**Use Approach 1 (DatabaseContext registration)** with these additions:
1. Create a unique temporary directory per merge job
2. Write the CSV data file into that directory
3. Register the data source with a unique name (e.g., UUID-based)
4. Execute the merge
5. Unregister the data source and clean up temp files
6. Wrap in try/finally to ensure cleanup

---

## 4. Complete Python UNO Mail Merge Example

```python
import uno
import os
import uuid
import tempfile
import csv
from com.sun.star.beans import PropertyValue

def connect_to_libreoffice(host="localhost", port=2002):
    """Connect to a running LibreOffice instance via UNO socket."""
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext
    )
    ctx = resolver.resolve(
        f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    )
    return ctx

def do_mail_merge(ctx, template_path, csv_data, output_dir, output_format="pdf"):
    """
    Execute a mail merge.
    
    Args:
        ctx: UNO component context
        template_path: Path to .docx template
        csv_data: List of dicts (each dict = one record)
        output_dir: Directory for output files
        output_format: "pdf", "docx", or "odt"
    """
    smgr = ctx.ServiceManager
    ds_name = f"lola_merge_{uuid.uuid4().hex[:8]}"
    temp_dir = tempfile.mkdtemp(prefix="lola_")
    
    try:
        # 1. Write CSV file
        if not csv_data:
            raise ValueError("No data provided")
        
        csv_path = os.path.join(temp_dir, "data.csv")
        fieldnames = list(csv_data[0].keys())
        
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_data)
        
        # 2. Register data source
        dbContext = smgr.createInstanceWithContext(
            "com.sun.star.sdb.DatabaseContext", ctx
        )
        ds = dbContext.createInstance()
        ds.URL = "sdbc:flat:" + uno.systemPathToFileUrl(temp_dir)
        
        info = []
        for name, value in [
            ("Extension", "csv"),
            ("HeaderLine", True),
            ("FieldDelimiter", ","),
            ("StringDelimiter", '"'),
            ("CharSet", "UTF-8"),
        ]:
            p = PropertyValue()
            p.Name = name
            p.Value = value
            info.append(p)
        
        ds.Info = tuple(info)
        dbContext.registerObject(ds_name, ds)
        ds.DatabaseDocument.store()
        
        # 3. Configure mail merge
        oMailMerge = smgr.createInstanceWithContext(
            "com.sun.star.text.MailMerge", ctx
        )
        
        oMailMerge.DocumentURL = uno.systemPathToFileUrl(
            os.path.abspath(template_path)
        )
        oMailMerge.DataSourceName = ds_name
        oMailMerge.CommandType = 0  # TABLE
        oMailMerge.Command = "data"  # CSV filename without extension
        oMailMerge.OutputType = 2   # FILE
        oMailMerge.OutputURL = uno.systemPathToFileUrl(
            os.path.abspath(output_dir)
        )
        oMailMerge.SaveAsSingleFile = False
        
        # Set output format
        filter_map = {
            "pdf": "writer_pdf_Export",
            "docx": "MS Word 2007 XML",
            "odt": "writer8",
        }
        oMailMerge.SaveFilter = filter_map.get(output_format, "writer_pdf_Export")
        
        # 4. Execute
        oMailMerge.execute(())
        
    finally:
        # 5. Cleanup: unregister data source
        try:
            dbContext.revokeObject(ds_name)
        except Exception:
            pass
        
        # Clean up temp files
        try:
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass
```

---

## 5. Running LibreOffice in Headless Server Mode

### Starting the Listener

```bash
soffice --headless --norestore --nologo \
  --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
```

**Flags explained:**
- `--headless` — no GUI (required for server use)
- `--norestore` — don't show recovery dialog on crash
- `--nologo` — skip splash screen
- `--accept=...` — open UNO socket listener

### Connection Pattern

```python
import uno

def get_uno_context(host="localhost", port=2002):
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext
    )
    try:
        ctx = resolver.resolve(
            f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
        )
        return ctx
    except Exception as e:
        raise ConnectionError(f"Cannot connect to LibreOffice: {e}")
```

### Connection Strategies

**Option A: Persistent connection (recommended for Lola)**
- Start LibreOffice once at container startup
- Keep a single connection open
- Reconnect on failure with exponential backoff
- Use a connection pool if concurrent requests needed (but LO is single-threaded)

**Option B: Start-per-request**
- Start LibreOffice for each request
- Slower (5-10 second startup overhead)
- More reliable (clean state each time)
- Not practical for production use

**Option C: Subprocess mode (avoid UNO socket)**
- Use `soffice --headless --convert-to pdf` for simple conversions
- Doesn't support mail merge (no programmatic control)
- Good for simple DOCX→PDF only

### Python-UNO Module Path

The `uno` Python module ships with LibreOffice. It's NOT a pip package. You need LibreOffice's Python or your system Python with the UNO path added:

```python
# In Docker, typically:
# /usr/lib/libreoffice/program/ contains uno.py
# The system python3 package python3-uno provides the bridge

import sys
sys.path.append('/usr/lib/libreoffice/program')
```

On Debian/Ubuntu Docker images:
```bash
apt-get install -y libreoffice python3-uno
```

---

## 6. Known Gotchas and Failure Modes

### LibreOffice Process Issues

1. **Single-threaded processing** — LibreOffice UNO is fundamentally single-threaded. Concurrent merge requests must be serialized (use a queue/lock in the Python service).

2. **Process crashes** — Malformed templates can crash LibreOffice entirely. The Python service must detect this (socket connection lost) and restart the LO process.

3. **Memory leaks** — Long-running LibreOffice instances accumulate memory. Consider periodic restarts (every N merges or every hour).

4. **Zombie processes** — If the Python process dies without cleaning up, LibreOffice continues running. Use process groups or a supervisor.

5. **Lock files** — LibreOffice creates `.~lock.*` files. In a container, the user profile directory should be isolated per instance.

6. **User profile corruption** — The LO user profile (`~/.config/libreoffice/`) can become corrupted, causing startup failures. Use `--env:UserInstallation=file:///tmp/lo_profile` to use a disposable profile.

### UNO Connection Issues

1. **Connection drops** — The UNO bridge can silently disconnect. Always wrap operations in try/except and implement reconnection.

2. **execute() argument** — Must pass a tuple `(())`, not a list `([])`. Lists cause `RuntimeException`.

3. **File URLs** — All paths must be `file:///` URLs, not filesystem paths. Use `uno.systemPathToFileUrl()`.

4. **Forward slashes only** — Even on systems that use backslashes, UNO requires forward slashes in file URLs.

### Mail Merge Specific Issues

1. **Field codes not flattened** — When using `SaveFilter = "writer8"` (ODF output), the output may still contain field codes instead of substituted values. Using `"writer_pdf_Export"` or `"MS Word 2007 XML"` typically flattens fields.

2. **Hidden paragraphs not hidden** — UNO mail merge may not evaluate conditional hidden text the same way the GUI wizard does. This is a known limitation.

3. **Data source registration persistence** — Registered data sources persist in LO's configuration. Always clean up with `revokeObject()`. Failure to do so fills up the registry.

4. **CSV encoding** — LibreOffice can be finicky about CSV encoding. Always use UTF-8 and explicitly set CharSet in the data source info.

5. **Empty fields** — Fields with no data may render as the field name (e.g., `«FieldName»`) instead of blank. Handle by ensuring all fields have at least empty string values.

6. **Filename collisions** — When `SaveAsSingleFile = False`, output files are named with a numeric suffix. Set `FileNamePrefix` to control naming.

---

## 7. DOCX ↔ LibreOffice Compatibility

### What Works Well

- ✅ Simple MERGEFIELD codes — field names preserved correctly
- ✅ Basic formatting (bold, italic, font size) on merged content
- ✅ `\* MERGEFORMAT` switch — formatting generally preserved
- ✅ Simple `IF MERGEFIELD = "value"` conditionals
- ✅ Tables with merge fields
- ✅ Headers/footers with merge fields
- ✅ Images in templates (static images preserved)

### What Breaks or Has Issues

- ⚠️ **Microsoft-only fonts** (Calibri, Cambria, etc.) — substituted with Liberation/DejaVu fonts. Install `ttf-mscorefonts-installer` or embed fonts.
- ⚠️ **Complex nested conditionals** — deeply nested `IF` within `IF` fields may not evaluate correctly
- ⚠️ **NEXT / NEXTIF fields** — work for simple cases but can fail with complex record-skipping logic
- ⚠️ **SmartArt, ActiveX controls, macros** — stripped or ignored
- ⚠️ **Custom XML data binding** — content controls bound to custom XML parts may lose their binding
- ❌ **VBA macros** — completely ignored
- ❌ **Word-specific field codes** — `DOCPROPERTY`, `STYLEREF`, `TC` (table of contents entries) have mixed support
- ❌ **Advanced conditional formatting** — conditional formatting rules in table cells may not transfer

### Font Handling in Docker

```dockerfile
# Install Microsoft core fonts (Arial, Times New Roman, etc.)
RUN apt-get update && apt-get install -y \
    ttf-mscorefonts-installer \
    fonts-liberation \
    fonts-dejavu \
    fontconfig \
    && fc-cache -f -v
```

### Line Spacing / Page Breaks

LibreOffice's line height and word wrapping calculations differ slightly from Word's. Documents may have different page break positions. For mail merge output, this is usually acceptable since the output is PDF (fixed layout).

---

## 8. Alternative Approaches Considered

### Why Not python-docx / docxtpl?

These libraries operate on the raw XML and support **template tags** (like `{{ field_name }}` in Jinja2). They do NOT understand Word's native MERGEFIELD codes:

- `python-docx` — can read/write paragraphs and runs but has no field code support
- `docxtpl` — uses Jinja2 templating in the XML, requires templates designed for it
- `docx-mailmerge` (PyPI) — actually does parse MERGEFIELD codes! But it does simple text substitution without LibreOffice's rendering engine, so formatting, conditionals, and NEXT fields don't work properly

### Why Not `soffice --convert-to`?

The command-line conversion (`soffice --headless --convert-to pdf file.docx`) is great for simple format conversion but doesn't support mail merge at all — it just converts the template as-is, with field codes showing as `«FieldName»`.

### Why UNO Is the Right Choice

LibreOffice's UNO mail merge uses the **same engine as the GUI wizard**. It properly:
- Evaluates all field codes
- Handles conditionals
- Processes NEXT records
- Applies formatting switches
- Renders with LibreOffice's full layout engine

The trade-off is complexity: UNO is harder to set up and more fragile than simple library calls. That's exactly why Lola wraps it in a managed service.

---

## References

- [LibreOffice MailMerge Service IDL](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1MailMerge.html)
- [MailMergeType Constants](https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1text_1_1MailMergeType.html)
- [DataAccessDescriptor](https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1sdb_1_1DataAccessDescriptor.html)
- [OpenOffice Forum: Python UNO Mail Merge (Solved)](https://forum.openoffice.org/en/forum/viewtopic.php?t=96001)
- [Stack Overflow: Flatten mailmerged documents](https://stackoverflow.com/questions/53441262/)
- [Office Open XML Field Codes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a845856b-6e8e-4c60-a5a3-5036186ebf1d)
- [LibreOffice Writer Guide: Mail Merge](https://books.libreoffice.org/en/WG71/WG7114-MailMerge.html)
