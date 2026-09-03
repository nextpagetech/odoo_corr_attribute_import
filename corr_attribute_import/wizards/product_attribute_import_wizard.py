import base64
import io

from odoo import _, api, fields, models
from odoo.exceptions import UserError


def _norm(s):
    if s is None:
        return ""
    return str(s).strip()


class AttributeImportWizard(models.TransientModel):
    _name = "attribute.import.wizard"
    _description = "Import Product Attributes and Values (no technical IDs)"

    file = fields.Binary(string="File", required=True, help="XLSX or CSV")
    file_name = fields.Char(string="Filename")
    mode = fields.Selection(
        [
            ("both", "Attributes and Values"),
            ("attributes", "Attributes only"),
            ("values", "Values only"),
        ],
        default="both",
        required=True,
    )
    update_existing = fields.Boolean(
        string="Update existing records",
        default=True,
        help="If enabled, existing attributes/values matched by name will be updated.",
    )
    create_missing_attribute = fields.Boolean(
        string="Create missing attributes when importing values",
        default=True,
    )
    log = fields.Text(readonly=True)

    # Accepted headers (case-insensitive)
    ATTR_HEADERS = {
        "name": {"name", "attribute", "attribute name"},
        "create_variant": {"create_variant", "create variants", "variants"},
        "sequence": {"sequence", "seq", "order"},
    }

    VAL_HEADERS = {
        "attribute": {"attribute", "attribute name", "attr", "product attribute"},
        "name": {"name", "value", "value name"},
        "sequence": {"sequence", "seq", "order"},
        # Custom field on product.attribute.value coming from
        # product_attribute_value_extras module
        "x_weight": {"x_weight", "weight"},
    }

    CREATE_VARIANT_MAP = {
        # normalized lowercase -> odoo value
        "always": "always",
        "no_variant": "no_variant",
        "no variant": "no_variant",
        "novariant": "no_variant",
        "dynamic": "dynamic",
    }

    def _is_xlsx(self):
        return (self.file_name or "").lower().endswith(".xlsx")

    def _is_csv(self):
        return (self.file_name or "").lower().endswith((".csv", ".txt"))

    def _read_xlsx(self):
        try:
            import openpyxl  # noqa: F401
        except Exception as e:
            raise UserError(
                _(
                    "openpyxl is required to read XLSX files. Please install it on the server.\nError: %s"
                )
                % e
            )

        data = base64.b64decode(self.file or b"")
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        result = {}

        def sheet_to_rows(ws):
            rows = []
            headers = []
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    headers = [(_ or "").strip() for _ in row]
                    continue
                rows.append({headers[j] if j < len(headers) else f"col{j}": v for j, v in enumerate(row)})
            return headers, rows

        for ws in wb.worksheets:
            headers, rows = sheet_to_rows(ws)
            result[ws.title.strip().lower()] = (headers, rows)
        return result

    def _read_csv(self):
        import csv

        data = base64.b64decode(self.file or b"")
        text = data.decode("utf-8-sig")

        # Single-sheet CSV; caller must indicate mode
        reader = csv.DictReader(io.StringIO(text))
        headers = reader.fieldnames or []
        rows = list(reader)
        return {"csv": (headers, rows)}

    @staticmethod
    def _match_map(headers, accepted):
        idx = {}
        lower = [(_ or "").strip().lower() for _ in headers]
        for key, aliases in accepted.items():
            for alias in aliases:
                if alias in lower:
                    idx[key] = lower.index(alias)
                    break
        return idx

    def _coerce_create_variant(self, val, default="always"):
        v = _norm(str(val)).lower()
        return self.CREATE_VARIANT_MAP.get(v, default)

    def _parse_int(self, v, default=10):
        try:
            if v is None or v == "":
                return default
            return int(float(v))
        except Exception:
            return default

    def _import_attributes_rows(self, headers, rows):
        ProductAttribute = self.env["product.attribute"].sudo()
        idx = self._match_map(headers, self.ATTR_HEADERS)
        if "name" not in idx:
            raise UserError(_("Attributes sheet: a 'Name' column is required."))

        created = updated = skipped = 0
        logs = []
        for r in rows:
            name = _norm(r.get(headers[idx["name"]], ""))
            if not name:
                skipped += 1
                continue
            create_variant = self._coerce_create_variant(
                r.get(headers[idx.get("create_variant", -1)], "") if idx.get("create_variant") is not None else ""
            )
            sequence = self._parse_int(r.get(headers[idx.get("sequence", -1)], ""))

            attr = ProductAttribute.search([("name", "=", name)], limit=1)
            vals = {"name": name, "create_variant": create_variant, "sequence": sequence}
            if attr:
                if self.update_existing:
                    attr.write(vals)
                    updated += 1
                    logs.append(f"Updated attribute: {name}")
                else:
                    skipped += 1
            else:
                ProductAttribute.create(vals)
                created += 1
                logs.append(f"Created attribute: {name}")

        return created, updated, skipped, logs

    def _import_values_rows(self, headers, rows):
        ProductAttribute = self.env["product.attribute"].sudo()
        ProductAttributeValue = self.env["product.attribute.value"].sudo()
        idx = self._match_map(headers, self.VAL_HEADERS)
        missing = {k for k in ("attribute", "name") if k not in idx}
        if missing:
            raise UserError(_("Values sheet: required columns missing: %s") % ", ".join(sorted(missing)))

        created = updated = skipped = 0
        logs = []
        for r in rows:
            attr_name = _norm(r.get(headers[idx["attribute"]], ""))
            val_name = _norm(r.get(headers[idx["name"]], ""))
            if not attr_name or not val_name:
                skipped += 1
                continue
            seq = self._parse_int(r.get(headers[idx.get("sequence", -1)], ""))

            attr = ProductAttribute.search([("name", "=", attr_name)], limit=1)
            if not attr:
                if not self.create_missing_attribute:
                    skipped += 1
                    logs.append(f"Skipped value '{val_name}' — attribute '{attr_name}' not found")
                    continue
                attr = ProductAttribute.create({
                    "name": attr_name,
                    "create_variant": "always",
                })
                logs.append(f"Created missing attribute: {attr_name}")

            pav = ProductAttributeValue.search([
                ("attribute_id", "=", attr.id),
                ("name", "=", val_name),
            ], limit=1)

            vals = {"name": val_name, "attribute_id": attr.id, "sequence": seq}

            # Optional custom weight column
            weight_pos = idx.get("x_weight")
            if weight_pos is not None and weight_pos < len(headers):
                raw_weight = r.get(headers[weight_pos], "")
                if raw_weight not in (None, ""):
                    try:
                        vals["x_weight"] = float(raw_weight)
                    except Exception:
                        # Ignore invalid weight values, continue with other data
                        pass
            if pav:
                if self.update_existing:
                    pav.write(vals)
                    updated += 1
                    logs.append(f"Updated value: {attr_name} / {val_name}")
                else:
                    skipped += 1
            else:
                ProductAttributeValue.create(vals)
                created += 1
                logs.append(f"Created value: {attr_name} / {val_name}")

        return created, updated, skipped, logs

    def action_import(self):
        self.ensure_one()
        if not self.file:
            raise UserError(_("Please upload a file."))

        sheets = {}
        if self._is_xlsx():
            sheets = self._read_xlsx()
        elif self._is_csv():
            sheets = self._read_csv()
        else:
            raise UserError(_("Unsupported file type. Please upload XLSX or CSV."))

        total_logs = []
        totals = {"created": 0, "updated": 0, "skipped": 0}

        def handle(name_key, handler, required=False):
            # Find the sheet by friendly names
            candidates = [
                k for k in sheets.keys()
                if k in {name_key, name_key.replace(" ", ""), name_key.lower()}
                   or any(name_key in k for k in sheets.keys())
            ]
            chosen = None
            for k in sheets.keys():
                if k in {name_key, name_key.lower(), name_key.replace(" ", "")}:
                    chosen = k
                    break
            if not chosen and candidates:
                chosen = candidates[0]

            if not chosen:
                if required:
                    raise UserError(_("Could not find sheet for '%s'.") % name_key)
                return

            headers, rows = sheets[chosen]
            c, u, s, logs = handler(headers, rows)
            totals["created"] += c
            totals["updated"] += u
            totals["skipped"] += s
            total_logs.extend(logs)

        # Routing by mode
        if self.mode in ("both", "attributes"):
            # Try sheets named 'Attributes' or similar; for CSV, the only sheet key is 'csv'
            key = "attributes" if self._is_xlsx() else list(sheets.keys())[0]
            handle(key, self._import_attributes_rows, required=(self.mode == "attributes"))

        if self.mode in ("both", "values"):
            key = "attribute values" if self._is_xlsx() else list(sheets.keys())[0]
            handle(key, self._import_values_rows, required=(self.mode == "values"))

        summary = (
            f"Created: {totals['created']}, Updated: {totals['updated']}, Skipped: {totals['skipped']}\n" +
            "\n".join(total_logs)
        )
        self.log = summary

        return {
            "type": "ir.actions.act_window",
            "res_model": self._name,
            "view_mode": "form",
            "res_id": self.id,
            "target": "new",
        }
