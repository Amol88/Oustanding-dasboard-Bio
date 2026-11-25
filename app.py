# Append slicers into existing chart drawing.xml OR create a new drawing.xml
        if existing_draw_target:
            draw_path = existing_draw_target.replace("../", "xl/").lstrip("/")
            if not draw_path.startswith("xl/"):
                draw_path = "xl/" + draw_path
            existing_draw_xml = files[draw_path].decode("utf-8")

            # openpyxl <wsDr> use karta hai (bina xdr: prefix ke) — namespaces add karein
            root_tag_end = existing_draw_xml.find(">")
            root_tag = existing_draw_xml[:root_tag_end]
            extra_ns = ""
            if "xmlns:xdr=" not in root_tag:
                extra_ns += ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
            if "xmlns:a=" not in root_tag:
                extra_ns += ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            if extra_ns:
                existing_draw_xml = root_tag + extra_ns + existing_draw_xml[root_tag_end:]

            if "</xdr:wsDr>" in existing_draw_xml:
                updated_draw_xml = existing_draw_xml.replace(
                    "</xdr:wsDr>", "".join(drawing_xml_items) + "</xdr:wsDr>", 1
                )
            else:
                updated_draw_xml = existing_draw_xml.replace(
                    "</wsDr>", "".join(drawing_xml_items) + "</wsDr>", 1
                )

            files[draw_path] = updated_draw_xml.encode("utf-8")
            need_drawing_tag = False
