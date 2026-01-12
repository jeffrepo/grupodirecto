# -*- coding: utf-8 -*-
from odoo import api, fields, models, _
from odoo.exceptions import UserError

import base64
import io

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class GdStockPorColorWizard(models.TransientModel):
    _name = "gd.stock.por.img.wizard"
    _description = "Stock por img / Stock por Lote (por Proveedor)"

    company_id = fields.Many2one(
        "res.company",
        string="Compañía",
        required=True,
        default=lambda self: self.env.company,
    )
    supplier_id = fields.Many2one(
        "res.partner",
        string="Proveedor",
        required=True,
    )

    file_data = fields.Binary(readonly=True)
    file_name = fields.Char(readonly=True)

    def _get_available_suppliers_domain(self):
        self.ensure_one()

        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("company_id", "in", [False, self.company_id.id]),
            ("partner_id.active", "=", True),
        ])

        partner_ids = supplierinfos.mapped("partner_id").ids
        return [("id", "in", partner_ids)]


    @api.onchange("company_id")
    def _onchange_company_id(self):
        for w in self:
            return {
                "domain": {
                    "supplier_id": w._get_available_suppliers_domain()
                }
            }

    # ----------------------------
    # Productos por proveedor
    # ----------------------------
    def _get_products_for_supplier(self):
        self.ensure_one()

        seller_domain = [
            ("partner_id", "=", self.supplier_id.id),
            ("company_id", "in", [False, self.company_id.id]),
        ]
        sellers = self.env["product.supplierinfo"].sudo().search(seller_domain)

        if not sellers:
            return self.env["product.product"]

        tmpl_ids = set(sellers.mapped("product_tmpl_id").ids)
        variant_ids = set(sellers.mapped("product_id").ids)

        # Traemos variantes:
        # - Las definidas directo en supplierinfo.product_id
        # - Todas las variantes de las plantillas supplierinfo.product_tmpl_id
        domain = ["|", ("id", "in", list(variant_ids)), ("product_tmpl_id", "in", list(tmpl_ids))]
        products = self.env["product.product"].sudo().search(domain)

        # Odoo 18: usa is_storable para filtrar inventariable/consumible según aplique
        products = products.filtered(lambda p: getattr(p, "is_storable", False))

        return products

    # ----------------------------
    # Stock actual por lote (stock.quant)
    # ----------------------------
    def _get_stock_by_lot(self, products):
        """Retorna:
        {
          product_id: [(lot_name_or_empty, qty), ...] ordenado por lot_name
        }
        """
        self.ensure_one()
        if not products:
            return {}

        domain = [
            ("product_id", "in", products.ids),
            ("location_id.usage", "=", "internal"),
            ("company_id", "in", [False, self.company_id.id]),
        ]

        groups = self.env["stock.quant"].sudo().read_group(
            domain,
            ["quantity:sum"],
            ["product_id", "lot_id"],
            lazy=False,
        )

        res = {}
        for g in groups:
            prod = g.get("product_id")
            if not prod:
                continue
            product_id = prod[0]
            lot = g.get("lot_id")  # (id, name) o False
            lot_name = lot[1] if lot else ""
            qty = g.get("quantity", 0.0) or 0.0

            # Si quieres excluir ceros, descomenta:
            # if abs(qty) < 1e-12:
            #     continue

            res.setdefault(product_id, []).append((str(lot_name), qty))

        # Orden por nombre de lote (como el Excel)
        for pid in res:
            res[pid].sort(key=lambda x: (x[0] or "").upper())

        return res

    # ----------------------------
    # Imagen (debajo del código)
    # ----------------------------
    def _prepare_image_bytesio(self, product, max_px=70):
        """Devuelve (bio, width_px, height_px) o (None, None, None)."""
        img_b64 = product.image_256 or product.image_128 or product.image_1920
        if not img_b64:
            return None, None, None

        try:
            raw = base64.b64decode(img_b64)
        except Exception:
            return None, None, None

        # Intentamos normalizar a PNG y reducir
        try:
            from PIL import Image  # pillow (normalmente ya viene por Odoo)

            im = Image.open(io.BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA")
            im.thumbnail((max_px, max_px))
            bio = io.BytesIO()
            im.save(bio, format="PNG")
            bio.seek(0)
            w, h = im.size
            return bio, w, h
        except Exception:
            # Fallback: devolvemos bytes tal cual (si xlsxwriter los acepta)
            bio = io.BytesIO(raw)
            return bio, max_px, max_px

    # ----------------------------
    # Generación Excel
    # ----------------------------
    def action_download_excel(self):
        self.ensure_one()

        if not xlsxwriter:
            raise UserError(_("No está instalado xlsxwriter en el entorno."))

        products = self._get_products_for_supplier()
        if not products:
            raise UserError(_("No se encontraron productos para el proveedor seleccionado."))

        stock_map = self._get_stock_by_lot(products)

        # Orden de productos como se espera (por referencia interna)
        products = products.sorted(key=lambda p: (p.default_code or "", p.id))

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True})
        ws = wb.add_worksheet("Stock por Color")

        # Column widths (según tu Excel)
        ws.set_column("A:A", 7.43)
        ws.set_column("B:B", 28)
        ws.set_column("C:C", 11.43)
        ws.set_column("D:D", 16.71)
        ws.set_column("E:E", 17.85)
        ws.set_column("F:F", 11.43)
        ws.set_column("G:G", 16)

        # Formats
        fmt_normal = wb.add_format({"font_name": "Calibri", "font_size": 11})
        fmt_bold = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True})

        fmt_header = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "bold": True,
            "bg_color": "#DAE3F3",
            "align": "center",
            "valign": "vcenter",
            "top": 1,          # thin
            "bottom": 2,       # double (xlsxwriter = 2)
        })
        fmt_header_left = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "bold": True,
            "bg_color": "#DAE3F3",
            "align": "left",
            "valign": "vcenter",
            "top": 1,
            "bottom": 2,
        })

        fmt_int_center_bold = wb.add_format({"bold": True, "align": "center"})
        fmt_code_bold = wb.add_format({"bold": True})
        fmt_desc_bold = wb.add_format({"bold": True})
        fmt_center = wb.add_format({"align": "center"})
        fmt_qty = wb.add_format({"num_format": "#,##0.00"})
        fmt_qty_bold = wb.add_format({"num_format": "#,##0.00", "bold": True, "bottom": 1})
        fmt_subt = wb.add_format({"bold": True, "bottom": 1})
        fmt_total_lbl = wb.add_format({"bold": True, "bg_color": "#F8CBAD", "bottom": 1})
        fmt_total_qty = wb.add_format({"bold": True, "bg_color": "#F8CBAD", "bottom": 1, "num_format": "#,##0.00"})

        # Header (como el layout)
        now_local = fields.Datetime.context_timestamp(self, fields.Datetime.now())
        date_str = now_local.strftime("%d/%m/%Y")
        time_str = now_local.strftime("%H:%M")

        # Filas 1-6 (0-index: 0-5)
        ws.write(0, 0, "Profit Plus Administrativo", fmt_normal)
        ws.write(0, 5, date_str, fmt_center)

        ws.write(1, 0, (self.company_id.name or "").upper(), fmt_normal)
        ws.write(1, 5, time_str, fmt_center)

        telefono = self.company_id.phone or ""
        nit = self.company_id.vat or ""

        ws.write(2, 0, "TEL.:", fmt_normal)
        ws.write(2, 1, telefono, fmt_normal)
        ws.write(3, 0, "N.I.T..:", fmt_normal)
        ws.write(3, 1, nit, fmt_normal)

        ws.write(4, 0, "ARTÍCULOS CON SU STOCK X LOTE", fmt_bold)

        supplier_name = self.supplier_id.display_name or self.supplier_id.name or ""
        ws.write(5, 0, f"Rangos: Proveedor: {supplier_name}", fmt_bold)

        # Encabezados (fila 8 en Excel => índice 7)
        ws.set_row(7, 26.25)
        headers = ["No.", "CODIGO", "MODELO", "DESCRIPCION", "NRO LOTE", "UNIDAD", "STOCK ACTUAL"]
        ws.write(7, 0, headers[0], fmt_header)
        ws.write(7, 1, headers[1], fmt_header)
        ws.write(7, 2, headers[2], fmt_header)
        ws.write(7, 3, headers[3], fmt_header)
        ws.write(7, 4, headers[4], fmt_header_left)
        ws.write(7, 5, headers[5], fmt_header)
        ws.write(7, 6, headers[6], fmt_header)

        # Línea separadora (fila 9 en Excel => índice 8)
        ws.set_row(8, 6)

        # Data inicia en fila 10 Excel => índice 9
        row = 9
        item_no = 0
        grand_total = 0.0
        PRODUCT_ROW_HEIGHT = 100

        for p in products:
            lots = stock_map.get(p.id, [])
            if not lots:
                # Si no hay lotes, ponemos una línea sin lote con stock 0 (o podrías sumar quants sin lote)
                lots = [("", 0.0)]

            # Subtotal por producto
            subtotal = sum(q for _, q in lots)
            grand_total += subtotal

            # 1era línea del producto (incluye 1er lote)
            item_no += 1
            first_lot, first_qty = lots[0]

            img_bio, img_w, img_h = self._prepare_image_bytesio(p, max_px=70)
            ws.set_row(row, PRODUCT_ROW_HEIGHT)


            ws.write(row, 0, item_no, fmt_int_center_bold)
            ws.write(row, 1, p.default_code or "", fmt_code_bold)
            ws.write(row, 2, "", fmt_code_bold)  # MODELO (si luego lo tienes, aquí lo llenas)
            ws.write(row, 3, p.name or "", fmt_desc_bold)
            ws.write(row, 4, first_lot or "", fmt_normal)
            ws.write(row, 5, (p.uom_id.name or ""), fmt_center)
            ws.write_number(row, 6, float(first_qty or 0.0), fmt_qty)

            # Insertar imagen "debajo" del código (col B) dentro de la misma fila
            if img_bio:
                ws.insert_image(
                    row, 1, "product.png",
                    {"image_data": img_bio, "x_offset": 4, "y_offset": 28}
                )

            row += 1

            # Resto de lotes
            for lot_name, qty in lots[1:]:
                ws.set_row(row, 15.0)
                ws.write(row, 4, lot_name or "", fmt_normal)
                ws.write(row, 5, (p.uom_id.name or ""), fmt_center)
                ws.write_number(row, 6, float(qty or 0.0), fmt_qty)
                row += 1

            # Subtotales
            ws.set_row(row, 15.0)
            ws.write(row, 5, "Subtotales:", fmt_subt)
            ws.write_number(row, 6, float(subtotal), fmt_qty_bold)
            row += 1

            # Espacio entre productos
            ws.set_row(row, 6.0)
            row += 1

        # Totales
        ws.set_row(row, 15.0)
        ws.write(row, 5, "Totales:", fmt_total_lbl)
        ws.write_number(row, 6, float(grand_total), fmt_total_qty)

        wb.close()
        output.seek(0)

        filename = f"Stock_por_img_{supplier_name}_{date_str.replace('/','-')}.xlsx"
        self.write({
            "file_name": filename,
            "file_data": base64.b64encode(output.getvalue()),
        })

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model={self._name}&id={self.id}"
                   f"&field=file_data&filename_field=file_name&download=true",
            "target": "self",
        }
