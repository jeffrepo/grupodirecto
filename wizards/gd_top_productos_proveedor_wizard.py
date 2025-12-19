# -*- coding: utf-8 -*-
import base64
import io
from datetime import datetime
import logging

from odoo import api, fields, models, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class GdTopProductosProveedorWizard(models.TransientModel):
    _name = "gd.top.productos.proveedor.wizard"
    _description = "Artículos más/menos vendidos por proveedor (Excel)"

    company_id = fields.Many2one(
        "res.company",
        string="Compañía",
        default=lambda self: self.env.company,
        required=True,
    )

    supplier_id = fields.Many2one(
        "res.partner",
        string="Proveedor",
        required=True,
    )

    date_from = fields.Date(string="Fecha inicio", required=True)
    date_to = fields.Date(string="Fecha fin", required=True)

    limit_products = fields.Integer(string="Cantidad de productos", default=10, required=True)

    order_mode = fields.Selection(
        [
            ("top", "Más vendido"),
            ("bottom", "Menos vendido"),
        ],
        string="Orden",
        default="top",
        required=True,
    )

    archivo = fields.Binary(string="Archivo", readonly=True)
    archivo_nombre = fields.Char(string="Nombre archivo", readonly=True)

    # -------------------------
    # Proveedores disponibles (desde supplierinfo)
    # -------------------------
    def _get_available_suppliers_domain(self):
        """Devuelve dominio para mostrar SOLO proveedores presentes en product.supplierinfo."""
        self.ensure_one()
        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("company_id", "in", [False, self.company_id.id]),
            ("partner_id.active", "=", True),
        ])
        partner_ids = supplierinfos.mapped("partner_id").ids
        _logger.info("[GD_REPORT] Available suppliers from supplierinfo (company=%s): %s",
                     self.company_id.id, partner_ids)
        return [("id", "in", partner_ids)]

    @api.onchange("company_id")
    def _onchange_company_id(self):
        """Al cambiar compañía, filtra proveedores disponibles desde supplierinfo."""
        for w in self:
            return {"domain": {"supplier_id": w._get_available_suppliers_domain()}}

    # -------------------------
    # Helpers read_group (Odoo cambia nombres de keys)
    # -------------------------
    @staticmethod
    def _rg_sum(group_dict, field_name):
        """Obtiene el valor sumado de read_group de forma tolerante a versiones.
        Soporta keys tipo: quantity_sum / quantity, price_subtotal_sum / price_subtotal, etc.
        """
        # intentos directos
        for k in (f"{field_name}_sum", field_name):
            if k in group_dict:
                val = group_dict.get(k)
                return float(val or 0.0)
        # fallback: cualquier key que termine en _sum y empiece con el field
        for k in group_dict.keys():
            if k.startswith(field_name) and k.endswith("_sum"):
                return float(group_dict.get(k) or 0.0)
        return 0.0

    # -------------------------
    # Debug (temporal)
    # -------------------------
    def _debug_dump_moves_for_products(self, product_ids):
        """Debug forense: ORM vs SQL y muestra ejemplos reales."""
        self.ensure_one()
        product_ids = list(map(int, product_ids or []))

        _logger.info(
            "[GD_DEBUG] db=%s uid=%s company=%s allowed_company_ids=%s product_ids=%s",
            self.env.cr.dbname,
            self.env.uid,
            self.company_id.id,
            self.env.context.get("allowed_company_ids"),
            product_ids,
        )

        aml_model = self.env["account.move.line"].sudo()

        c1 = aml_model.search_count([("product_id", "in", product_ids)])
        _logger.info("[GD_DEBUG] ORM count product_id in list: %s", c1)

       
        c_ok = aml_model.search_count([
            ("product_id", "in", product_ids),
            ("display_type", "=", "product"),
        ])
        _logger.info("[GD_DEBUG] ORM count display_type='product': %s", c_ok)

        sample = aml_model.search(
            [
                ("product_id", "in", product_ids),
                ("display_type", "=", "product"),
            ],
            limit=15
        )
        for line in sample:
            m = line.move_id
            _logger.info(
                "[GD_DEBUG] SAMPLE aml_id=%s move=%s(%s) type=%s state=%s date=%s invoice_date=%s product_id=%s qty=%s display_type=%s",
                line.id, m.name, m.id, m.move_type, m.state, m.date, m.invoice_date,
                line.product_id.id, line.quantity, line.display_type,
            )

        if product_ids:
            self.env.cr.execute(
                "SELECT id, product_id, move_id FROM account_move_line WHERE product_id IN %s LIMIT 30",
                [tuple(product_ids)]
            )
            rows = self.env.cr.fetchall()
            _logger.info("[GD_DEBUG] SQL rows=%s sample=%s", len(rows), rows[:10])

    # -------------------------
    # Validaciones
    # -------------------------
    def _validate_params(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_("Falta la librería 'xlsxwriter' en tu entorno Python."))

        if self.date_from > self.date_to:
            raise UserError(_("La fecha inicio no puede ser mayor que la fecha fin."))

        if self.limit_products <= 0:
            raise UserError(_("La cantidad de productos debe ser mayor a 0."))

    # -------------------------
    # Proveedor -> Productos
    # -------------------------
    def _get_product_ids_for_supplier(self):
        self.ensure_one()

        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("partner_id", "=", self.supplier_id.id),
            ("company_id", "in", [False, self.company_id.id]),
        ])

        _logger.info("[GD_REPORT] supplierinfo found: %s", len(supplierinfos))

        product_ids = set()

        # 1) Supplierinfo a nivel VARIANTE
        variant_supplierinfos = supplierinfos.filtered(lambda s: s.product_id)
        for si in variant_supplierinfos:
            product_ids.add(si.product_id.id)
        _logger.info("[GD_REPORT] supplierinfo with product_id (variant-level): %s", len(variant_supplierinfos))

        # 2) Supplierinfo a nivel TEMPLATE 
        template_supplierinfos = supplierinfos.filtered(lambda s: not s.product_id)
        tmpl_ids = template_supplierinfos.mapped("product_tmpl_id").ids
        _logger.info("[GD_REPORT] supplierinfo template-level: %s (templates=%s)", len(template_supplierinfos), tmpl_ids)

        if tmpl_ids:
            variants = self.env["product.product"].sudo().search([
                ("product_tmpl_id", "in", tmpl_ids),
                ("active", "=", True),
            ])
            product_ids.update(variants.ids)

        _logger.info("[GD_REPORT] FINAL resolved product_ids: %s", list(product_ids))
        return list(product_ids)

    # -------------------------
    # Facturas -> agregar por producto (ventas cliente)
    # Ranking por CANTIDAD (lo que te pide el reporte "más/menos vendido")
    # -------------------------
    def _get_sales_by_product(self, product_ids):
        """Retorna lista de dicts: product_id, qty, amount (neto).
        - qty = cantidad neta vendida (ventas - devoluciones)
        - amount = monto neto (price_subtotal)
        """
        self.ensure_one()
        if not product_ids:
            return []

        aml = self.env["account.move.line"].sudo()

        
        base_domain = [
            ("product_id", "in", product_ids),
            ("display_type", "=", "product"),
            ("move_id.company_id", "=", self.company_id.id),
            ("move_id.state", "=", "posted"),
            ("move_id.date", ">=", self.date_from),
            ("move_id.date", "<=", self.date_to),
        ]

        inv_domain = base_domain + [("move_id.move_type", "=", "out_invoice")]
        ref_domain = base_domain + [("move_id.move_type", "=", "out_refund")]

        _logger.info("[GD_REPORT] Sales domain (out_invoice): %s", inv_domain)
        _logger.info("[GD_REPORT] Refund domain (out_refund): %s", ref_domain)
        _logger.info("[GD_REPORT] product_ids filter size: %s", len(product_ids))

        inv_groups = aml.read_group(
            inv_domain,
            ["product_id", "quantity:sum", "price_subtotal:sum"],
            ["product_id"],
            lazy=False,
        )
        ref_groups = aml.read_group(
            ref_domain,
            ["product_id", "quantity:sum", "price_subtotal:sum"],
            ["product_id"],
            lazy=False,
        )

        _logger.info("[GD_REPORT] out_invoice groups: %s", len(inv_groups))
        _logger.info("[GD_REPORT] out_refund groups: %s", len(ref_groups))
        _logger.info("[GD_DEBUG] inv_groups sample: %s", inv_groups[:5])

        by_product = {}

        # Ventas
        for g in inv_groups:
            pid = g["product_id"][0]
            qty = self._rg_sum(g, "quantity")
            amt = self._rg_sum(g, "price_subtotal")
            by_product[pid] = {"product_id": pid, "qty": qty, "amount": amt}

        # Devoluciones (neteo)
        for g in ref_groups:
            pid = g["product_id"][0]
            qty = self._rg_sum(g, "quantity")
            amt = self._rg_sum(g, "price_subtotal")

            if pid not in by_product:
                by_product[pid] = {"product_id": pid, "qty": 0.0, "amount": 0.0}

            # Si refund ya viene negativo, se suma; si viene positivo, se resta.
            if qty < 0 or amt < 0:
                by_product[pid]["qty"] += qty
                by_product[pid]["amount"] += amt
            else:
                by_product[pid]["qty"] -= qty
                by_product[pid]["amount"] -= amt

        rows = list(by_product.values())

        # Filtrar sin movimiento neto
        rows = [r for r in rows if (abs(r["qty"]) > 1e-9 or abs(r["amount"]) > 1e-9)]

        # Ranking por cantidad (más/menos vendido)
        reverse = (self.order_mode == "top")
        rows.sort(key=lambda r: (r["qty"], r["amount"]), reverse=reverse)

        return rows[: self.limit_products]

    # -------------------------
    # Excel
    # -------------------------
    def _build_xlsx(self, rows):
        self.ensure_one()

        output = io.BytesIO()
        sheet_name = "10 + Vendidos" if self.order_mode == "top" else "10 + Menos Vendidos"
        wb = xlsxwriter.Workbook(output, {"in_memory": True})
        ws = wb.add_worksheet(sheet_name)

        # Columnas
        ws.set_column("A:A", 6.0)
        ws.set_column("B:B", 17.43)
        ws.set_column("C:C", 21.57)
        ws.set_column("D:D", 37.43)
        ws.set_column("E:E", 10.57)
        ws.set_column("F:F", 11.43)

        # formato
        fmt_calibri = wb.add_format({"font_name": "Calibri", "font_size": 11})
        fmt_time = wb.add_format({"font_name": "Calibri", "font_size": 11, "num_format": "h:mm AM/PM"})

        fmt_header = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "top": 1,
            "bottom": 6,
            "bg_color": "#DAE3F3",
            "pattern": 1,
        })

        fmt_no_first = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "top": 6, "bottom": 1})
        fmt_no = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "bottom": 1})

        fmt_text_first = wb.add_format({"font_name": "Arial", "font_size": 10, "valign": "bottom", "top": 6, "bottom": 1})
        fmt_text = wb.add_format({"font_name": "Arial", "font_size": 10, "valign": "bottom", "bottom": 1})

        fmt_medida_first = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "top": 6, "bottom": 1})
        fmt_medida = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "bottom": 1})

        fmt_qty_first = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "top": 6, "bottom": 1, "num_format": "#,##0.00"})
        fmt_qty = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True, "align": "center", "valign": "bottom", "bottom": 1, "num_format": "#,##0.00"})

        fmt_total_label = wb.add_format({"font_name": "Calibri", "font_size": 11, "top": 1, "bottom": 6})
        fmt_total_value = wb.add_format({"font_name": "Calibri", "font_size": 11, "bold": True, "align": "center", "top": 1, "bottom": 6, "num_format": "#,##0.00"})

        
        for r in range(0, 7):
            ws.set_row(r, 15.0)
        ws.set_row(9, 30.75)

        
        now_local = fields.Datetime.context_timestamp(self, fields.Datetime.now())
        ws.write(0, 0, "Profit Plus Administrativo", fmt_calibri)
        ws.write(0, 5, now_local.strftime("%d/%m/%Y"), fmt_calibri)

        ws.write(1, 0, (self.company_id.name or "").upper(), fmt_calibri)

        excel_time = datetime(1900, 1, 1, now_local.hour, now_local.minute, 0)
        ws.write_datetime(1, 5, excel_time, fmt_time)

        # TEL y NIT de la compañía
        telefono = self.company_id.phone or ""
        nit = self.company_id.vat or ""

        ws.write(2, 0, "TEL.:", fmt_calibri)
        ws.write(2, 1, telefono, fmt_calibri)

        ws.write(3, 0, "N.I.T.:", fmt_calibri)
        ws.write(3, 1, nit, fmt_calibri)


        titulo = "Artículos con más Ventas (Orden: Cantidad)" if self.order_mode == "top" else "Artículos con menos Ventas (Orden: Cantidad)"
        ws.write(4, 0, titulo, fmt_calibri)

        proveedor_nombre = self.supplier_id.display_name or ""
        ws.write(
            5, 0,
            f"Rangos: Fecha: {self.date_from.strftime('%d/%m/%Y')} Hasta {self.date_to.strftime('%d/%m/%Y')}; Proveedor: {proveedor_nombre}; ",
            fmt_calibri
        )

        ws.write(6, 0, f"Los Mejores: {self.limit_products}", fmt_calibri)

        headers = ["No.", "ARTICULO", "MODELO", "DESCRIPCION", "MEDIDA", "CANTIDAD"]
        for col, h in enumerate(headers):
            ws.write(9, col, h, fmt_header)

        start_row = 10
        product_map = {p.id: p for p in self.env["product.product"].sudo().browse([r["product_id"] for r in rows])}

        current_row = start_row
        for i, r in enumerate(rows, start=1):
            ws.set_row(current_row, 15.0 if current_row != start_row else 15.75)

            is_first = (current_row == start_row)
            f_no = fmt_no_first if is_first else fmt_no
            f_txt = fmt_text_first if is_first else fmt_text
            f_med = fmt_medida_first if is_first else fmt_medida
            f_qty = fmt_qty_first if is_first else fmt_qty

            if i == 1:
                ws.write_number(current_row, 0, 1, f_no)
            else:
                ws.write_formula(current_row, 0, f"=+A{current_row}+1", f_no)

            prod = product_map.get(r["product_id"])

            articulo = (prod.default_code or "") if prod else ""
            modelo = (prod.product_tmpl_id.name or "") if prod and prod.product_tmpl_id else ""
            descripcion = (prod.name or "") if prod else ""
            medida = (prod.uom_id.name or "") if prod and prod.uom_id else ""

            ws.write(current_row, 1, articulo, f_txt)
            ws.write(current_row, 2, modelo, f_txt)
            ws.write(current_row, 3, descripcion, f_txt)
            ws.write(current_row, 4, medida, f_med)
            ws.write_number(current_row, 5, float(r["qty"]), f_qty)

            current_row += 1

      
        blank_row = current_row
        ws.set_row(blank_row, 6.0)
        current_row += 1

        # Totales
        totals_row = current_row
        ws.set_row(totals_row, 15.75)
        ws.write(totals_row, 4, "Totales:", fmt_total_label)

        first_excel_row = start_row + 1
        last_excel_row = (start_row + len(rows))
        ws.write_formula(totals_row, 5, f"=SUM(F{first_excel_row}:F{last_excel_row})", fmt_total_value)

        wb.close()
        output.seek(0)
        return output.getvalue()

    # -------------------------
    # Acción principal
    # -------------------------
    def action_download_excel(self):
        self.ensure_one()
        self._validate_params()

        _logger.info(
            "[GD_REPORT] Wizard run id=%s company=%s supplier_id=%s supplier=%s dates=%s..%s limit=%s order=%s",
            self.id, self.company_id.id, self.supplier_id.id, self.supplier_id.display_name,
            self.date_from, self.date_to, self.limit_products, self.order_mode
        )

        product_ids = self._get_product_ids_for_supplier()
        _logger.info("[GD_REPORT] supplier->products: %s products found", len(product_ids))

        if product_ids:
            sample = self.env["product.product"].sudo().browse(product_ids[:10]).mapped("display_name")
            _logger.info("[GD_REPORT] sample products: %s", sample)

        if not product_ids:
            raise UserError(_("No se encontraron productos vinculados a este proveedor en la pestaña Compras."))

        # Debug temporal
        self._debug_dump_moves_for_products(product_ids)

        rows = self._get_sales_by_product(product_ids)

        _logger.info("[GD_REPORT] aggregated rows: %s", len(rows))
        if rows:
            _logger.info("[GD_REPORT] top rows preview: %s", rows[:10])

        if not rows:
            raise UserError(_("No hay movimientos en el rango de fechas para este proveedor."))

        xlsx_content = self._build_xlsx(rows)
        filename = f"Reporte_Articulos_{self.supplier_id.ref or self.supplier_id.id}_{self.date_from}_{self.date_to}.xlsx"

        self.write({
            "archivo": base64.b64encode(xlsx_content),
            "archivo_nombre": filename,
        })

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model={self._name}&id={self.id}&field=archivo&filename_field=archivo_nombre&download=true",
            "target": "self",
        }
