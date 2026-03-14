"""
UNO mail merge implementation.

Follows the pattern in docs/UNO-MAILMERGE.md:
  1. Write CSV to temp directory
  2. Register as data source via DatabaseContext
  3. Configure MailMerge UNO object and execute
  4. Collect output file paths
  5. Clean up data source and temp files
"""

import csv
import logging
import os
import shutil
import tempfile
import uuid

logger = logging.getLogger(__name__)

FILTER_MAP = {
    "pdf": "writer_pdf_Export",
    "docx": "MS Word 2007 XML",
    "odt": "writer8",
}


def mail_merge(
    ctx,
    template_path: str,
    data: list[dict],
    output_dir: str,
    output_format: str = "pdf",
) -> list[str]:
    """
    Execute a UNO mail merge.

    Args:
        ctx: UNO component context (from UnoClient.ctx).
        template_path: Absolute path to .docx template.
        data: List of dicts — each dict is one record, keys are field names.
        output_dir: Absolute path to output directory (created if needed).
        output_format: "pdf", "docx", or "odt".

    Returns:
        List of absolute paths to output files.
    """
    import uno
    from com.sun.star.beans import PropertyValue

    if not data:
        raise ValueError("No data provided")

    os.makedirs(output_dir, exist_ok=True)

    save_filter = FILTER_MAP.get(output_format, "writer_pdf_Export")
    smgr = ctx.ServiceManager
    ds_name = f"lola_merge_{uuid.uuid4().hex[:8]}"
    temp_dir = tempfile.mkdtemp(prefix="lola_")

    db_context = None
    try:
        # 1. Write CSV file
        csv_path = os.path.join(temp_dir, "data.csv")
        fieldnames = list(data[0].keys())
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(data)

        # 2. Register CSV directory as data source
        db_context = smgr.createInstanceWithContext(
            "com.sun.star.sdb.DatabaseContext", ctx
        )
        ds = db_context.createInstance()
        ds.URL = "sdbc:flat:" + uno.systemPathToFileUrl(temp_dir)

        info_props = []
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
            info_props.append(p)

        ds.Info = tuple(info_props)
        db_context.registerObject(ds_name, ds)
        ds.DatabaseDocument.store()

        # 3. Configure MailMerge UNO object
        mail_merge_obj = smgr.createInstanceWithContext(
            "com.sun.star.text.MailMerge", ctx
        )
        mail_merge_obj.DocumentURL = uno.systemPathToFileUrl(os.path.abspath(template_path))
        mail_merge_obj.DataSourceName = ds_name
        mail_merge_obj.CommandType = 0   # TABLE
        mail_merge_obj.Command = "data"  # CSV filename without extension
        mail_merge_obj.OutputType = 2    # FILE
        mail_merge_obj.OutputURL = uno.systemPathToFileUrl(os.path.abspath(output_dir))
        mail_merge_obj.SaveAsSingleFile = False
        mail_merge_obj.SaveFilter = save_filter

        # 4. Execute merge — must pass tuple, not list
        mail_merge_obj.execute(())

        # 5. Collect output files
        ext_map = {"pdf": ".pdf", "docx": ".docx", "odt": ".odt"}
        ext = ext_map.get(output_format, ".pdf")
        output_files = sorted(
            os.path.join(output_dir, f)
            for f in os.listdir(output_dir)
            if f.endswith(ext)
        )
        return output_files

    finally:
        # Cleanup: unregister data source
        if db_context is not None:
            try:
                db_context.revokeObject(ds_name)
            except Exception:
                pass

        # Remove temp CSV directory
        shutil.rmtree(temp_dir, ignore_errors=True)
