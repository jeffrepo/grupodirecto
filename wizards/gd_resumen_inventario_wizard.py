# -*- coding: utf-8 -*-
import base64
import io
import logging
from datetime import datetime, time, timedelta

import pytz

from odoo import api, fields, models, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class GdResumenInventarioWizard(models.TransientModel):
    _name = "gd.resumen.inventario.wizard"
    _description = "Reporte 3 - Resumen de Inventario por Proveedor (Excel)"

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

    date_from = fields.Date(string="Desde", required=True)
    date_to = fields.Date(string="Hasta", required=True)

    archivo = fields.Binary(string="Archivo", readonly=True)
    archivo_nombre = fields.Char(string="Nombre archivo", readonly=True)

    # -------------------------
    # Helpers read_group (tolerante a variantes de keys)
    # -------------------------
    @staticmethod
    def _rg_sum(group_dict, field_name):
        """Tolera variantes de keys en read_group entre builds."""
        for k in (f"{field_name}_sum", field_name):
            if k in group_dict:
                return float(group_dict.get(k) or 0.0)
        for k in group_dict.keys():
            if k.startswith(field_name) and k.endswith("_sum"):
                return float(group_dict.get(k) or 0.0)
        return 0.0

    # -------------------------
    # Dominio de proveedores (IGUAL QUE REPORTE 2): desde product.supplierinfo
    # -------------------------
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
            return {"domain": {"supplier_id": w._get_available_suppliers_domain()}}

    # -------------------------
    # Validaciones
    # -------------------------
    def _validate_params(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_("Falta la librería 'xlsxwriter' en tu entorno Python."))

        if self.date_from > self.date_to:
            raise UserError(_("Rango inválido: 'Desde' no puede ser mayor que 'Hasta'."))

    # -------------------------
    # Proveedor -> Productos (IGUAL QUE REPORTE 2): supplierinfo variante/plantilla
    # -------------------------
    def _get_product_ids_for_supplier(self):
        self.ensure_one()

        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("partner_id", "=", self.supplier_id.id),
            ("company_id", "in", [False, self.company_id.id]),
        ])

        product_ids = set()

        # 1) supplierinfo a nivel variante
        for si in supplierinfos.filtered(lambda s: s.product_id):
            product_ids.add(si.product_id.id)

        # 2) supplierinfo a nivel plantilla => todas las variantes
        tmpl_ids = supplierinfos.filtered(lambda s: not s.product_id).mapped("product_tmpl_id").ids
        if tmpl_ids:
            variants = self.env["product.product"].sudo().search([
                ("product_tmpl_id", "in", tmpl_ids),
                ("active", "=", True),
            ])
            product_ids.update(variants.ids)

        # Filtrar solo stockeables/consumibles (evita servicios)
        if product_ids:
            Product = self.env["product.product"].sudo()
            type_field = "detailed_type" if "detailed_type" in Product._fields else "type"
            prods = Product.browse(list(product_ids)).filtered(
                lambda p: getattr(p, type_field) in ("product", "consu")
            )
            return prods.ids

        return []

    def _get_products_for_supplier(self):
        product_ids = self._get_product_ids_for_supplier()
        return self.env["product.product"].sudo().browse(product_ids).sorted(
            key=lambda p: (p.default_code or "", p.id)
        )

    # -------------------------
    # Rango datetime en UTC
    # -------------------------
    def _get_utc_range(self):
        self.ensure_one()
        user_tz = pytz.timezone(self.env.user.tz or "UTC")

        dt_from_local = user_tz.localize(datetime.combine(self.date_from, time.min))
        # time.max trae microsegundos; lo dejamos en 0 para evitar cosas raras en filtros
        dt_to_local = user_tz.localize(datetime.combine(self.date_to, time.max.replace(microsecond=0)))

        dt_from_utc = dt_from_local.astimezone(pytz.UTC)
        dt_to_utc = dt_to_local.astimezone(pytz.UTC)

        # stock inicial = justo antes de iniciar el rango
        dt_open_utc = (dt_from_local - timedelta(seconds=1)).astimezone(pytz.UTC)
        return dt_from_utc, dt_to_utc, dt_open_utc

    # -------------------------
    # Cantidad DONE en move lines (Odoo 18: qty_done puede NO ser store)
    # -------------------------
    def _get_move_line_qty_field(self, require_store=True):
        """
        Retorna el nombre del campo de cantidad en stock.move.line.

        - Si require_store=True: busca el primero que exista y sea store (para read_group).
        - Si require_store=False: busca el primero que exista aunque sea computed (para fallback sumado en python).
        """
        MoveLine = self.env["stock.move.line"]
        candidates = (
            "quantity",        # en algunas versiones es el store
            "qty_done",        # clásico, pero en tu build no es store
            "quantity_done",
            "done_qty",
        )
        for fname in candidates:
            f = MoveLine._fields.get(fname)
            if not f:
                continue
            if require_store and not getattr(f, "store", False):
                continue
            # debe ser numérico
            if getattr(f, "type", None) not in ("float", "integer", "monetary"):
                continue
            return fname
        return None

    def _sum_moves(self, products, dt_from_utc, dt_to_utc, src_usage, dest_usage):
        """Devuelve dict product_id -> qty (movimientos DONE por uso de ubicaciones)."""
        self.ensure_one()
        if not products:
            return {}

        MoveLine = self.env["stock.move.line"].sudo()

        domain = [
            ("state", "=", "done"),
            ("company_id", "=", self.company_id.id),
            ("product_id", "in", products.ids),
            ("date", ">=", fields.Datetime.to_string(dt_from_utc)),
            ("date", "<=", fields.Datetime.to_string(dt_to_utc)),
            ("location_id.usage", "=", src_usage),
            ("location_dest_id.usage", "=", dest_usage),
        ]

        # 1) Camino rápido: read_group con un campo store
        qty_field_store = self._get_move_line_qty_field(require_store=True)
        if qty_field_store:
            groups = MoveLine.read_group(
                domain,
                [f"{qty_field_store}:sum"],
                ["product_id"],
                lazy=False,
            )
            res = {}
            for g in groups:
                if g.get("product_id"):
                    pid = g["product_id"][0]
                    res[pid] = self._rg_sum(g, qty_field_store)
            return res

        # 2) Fallback: sumar en python (si no hay campo store usable)
        qty_field_any = self._get_move_line_qty_field(require_store=False) or "qty_done"
        _logger.warning(
            "[GD_R3] No hay campo qty store para read_group en stock.move.line. "
            "Usando fallback python con campo=%s (puede ser más lento).",
            qty_field_any,
        )

        res = {}
        for ml in MoveLine.search(domain):
            pid = ml.product_id.id
            qty = float(getattr(ml, qty_field_any, 0.0) or 0.0)
            res[pid] = res.get(pid, 0.0) + qty
        return res

    # -------------------------
    # Excel (maqueta del archivo que me pasaste)
    # -------------------------
    def _build_xlsx(self, lines):
        self.ensure_one()
        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True})
        ws = wb.add_worksheet("Movimientos de Inventario")

        # Column widths (según tu Excel)
        ws.set_column("A:A", 17.42578125)
        ws.set_column("B:B", 11.42578125)
        ws.set_column("C:C", 15.42578125)
        ws.set_column("D:D", 11.42578125)
        ws.set_column("E:E", 11.42578125)
        ws.set_column("F:F", 13.0)
        ws.set_column("G:G", 15.5703125)
        ws.set_column("H:H", 11.42578125)
        ws.set_column("I:I", 13.0)
        ws.set_column("J:J", 12.7109375)

        # Row heights
        ws.set_row(7, 26.25)  # header row (Excel row 8)

        # Formats
        fmt_base = wb.add_format({"font_name": "Arial", "font_size": 10})
        fmt_bold = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True})
        fmt_time = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "h:mm AM/PM"})
        fmt_num = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "#,##0.00"})

        fmt_header = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,     # IMPORTANTE por los "\n" del header
            "top": 1,              # thin
            "bottom": 6,           # double
            "bg_color": "#DAE3F3",
            "pattern": 1,
        })

        # Header text
        now_local = fields.Datetime.context_timestamp(self, fields.Datetime.now())
        ws.write(0, 0, "Profit Plus Administrativo", fmt_base)
        ws.write(0, 9, now_local.strftime("%d/%m/%Y"), fmt_base)

        ws.write(1, 0, (self.company_id.name or "").upper(), fmt_base)
        excel_time = datetime(1900, 1, 1, now_local.hour, now_local.minute, 0)
        ws.write_datetime(1, 9, excel_time, fmt_time)

        ws.write(2, 0, "TEL.:", fmt_base)
        ws.write(2, 1, self.company_id.phone or "", fmt_base)

        # OJO: en tu template viene como "N.I.T..:"
        ws.write(3, 0, "N.I.T..:", fmt_base)
        ws.write(3, 1, self.company_id.vat or "", fmt_base)

        ws.write(4, 0, "LIBRO DE INVENTARIO", fmt_base)

        proveedor_codigo = (self.supplier_id.ref or self.supplier_id.name or "").strip()
        ws.write(
            5, 0,
            f"Rangos: Fecha: {self.date_from.strftime('%d/%m/%Y')} Hasta {self.date_to.strftime('%d/%m/%Y')}; "
            f"Proveedor: {proveedor_codigo}",
            fmt_bold
        )

        # Table header (Excel row 8 => index 7)
        headers = [
            "ARTICULO", "MODELO", "DESCRIPCION", "UNIDAD",
            "STOCK \nINICIAL", "COMPRAS", "DEVOLUCIONES", "VENTAS", "SALIDA COJINES",
            "STOCK \nFINAL"
        ]
        for col, h in enumerate(headers):
            ws.write(7, col, h, fmt_header)

        # Data rows start at Excel row 9 => index 8
        start_row = 8
        for i, line in enumerate(lines):
            r = start_row + i
            ws.set_row(r, 16.5)

            ws.write(r, 0, line["articulo"], fmt_base)
            ws.write(r, 1, "", fmt_base)  # MODELO vacío como tu Excel
            ws.write(r, 2, line["descripcion"], fmt_base)
            ws.write(r, 3, line["unidad"], fmt_base)

            ws.write_number(r, 4, line["stock_inicial"], fmt_num)
            ws.write_number(r, 5, line["compras"], fmt_num)
            ws.write_number(r, 6, line["devoluciones"], fmt_num)
            ws.write_number(r, 7, line["ventas"], fmt_num)

            # SALIDA COJINES: aún sin regla -> queda en blanco
            ws.write(r, 8, "", fmt_base)

            # STOCK FINAL: =E - H - I + G + F  (idéntico a tu plantilla)
            excel_row = r + 1
            ws.write_formula(r, 9, f"=E{excel_row}-H{excel_row}-I{excel_row}+G{excel_row}+F{excel_row}", fmt_num)

        # Blank row (como tu template)
        blank_row = start_row + len(lines)
        ws.set_row(blank_row, 5.25)

        # Totals row
        total_row = blank_row + 1
        ws.set_row(total_row, 15.75)

        first = start_row + 1
        last = blank_row  # suma incluye la fila en blanco (no afecta, está vacía)

        ws.write(total_row, 0, "Totales:", fmt_bold)
        ws.write_formula(total_row, 4, f"=SUM(E{first}:E{last})", fmt_num)
        ws.write_formula(total_row, 5, f"=SUM(F{first}:F{last})", fmt_num)
        ws.write_formula(total_row, 6, f"=SUM(G{first}:G{last})", fmt_num)
        ws.write_formula(total_row, 7, f"=SUM(H{first}:H{last})", fmt_num)
        ws.write_formula(total_row, 8, f"=SUM(I{first}:I{last})", fmt_num)
        ws.write_formula(total_row, 9, f"=SUM(J{first}:J{last})", fmt_num)

        wb.close()
        output.seek(0)
        return output.getvalue()

    # -------------------------
    # Acción principal
    # -------------------------
    def action_download_excel(self):
        self.ensure_one()
        self._validate_params()

        products = self._get_products_for_supplier()
        if not products:
            raise UserError(_("No se encontraron productos vinculados a este proveedor en la pestaña Compras."))

        dt_from_utc, dt_to_utc, dt_open_utc = self._get_utc_range()

        compras = self._sum_moves(products, dt_from_utc, dt_to_utc, src_usage="supplier", dest_usage="internal")
        devoluciones = self._sum_moves(products, dt_from_utc, dt_to_utc, src_usage="customer", dest_usage="internal")
        ventas = self._sum_moves(products, dt_from_utc, dt_to_utc, src_usage="internal", dest_usage="customer")

        # Stock inicial (histórico)
        stock_inicial = {}
        to_date = fields.Datetime.to_string(dt_open_utc)
        prods_to_date = products.with_context(
            to_date=to_date,
            company_id=self.company_id.id,
            allowed_company_ids=[self.company_id.id],
        )
        for p in prods_to_date:
            stock_inicial[p.id] = float(p.qty_available or 0.0)

        lines = []
        for p in products:
            ini = float(stock_inicial.get(p.id, 0.0) or 0.0)
            com = float(compras.get(p.id, 0.0) or 0.0)
            dev = float(devoluciones.get(p.id, 0.0) or 0.0)
            ven = float(ventas.get(p.id, 0.0) or 0.0)

            # salida cojines: sin regla => lo dejamos vacío en Excel
            if ini == 0.0 and com == 0.0 and dev == 0.0 and ven == 0.0:
                continue

            lines.append({
                "articulo": p.default_code or "",
                "descripcion": p.name or p.display_name or "",
                "unidad": p.uom_id.name or "",
                "stock_inicial": ini,
                "compras": com,
                "devoluciones": dev,
                "ventas": ven,
            })

        if not lines:
            raise UserError(_("No hay movimientos/existencias en el rango para este proveedor."))

        xlsx_content = self._build_xlsx(lines)

        supplier_code = (self.supplier_id.ref or str(self.supplier_id.id) or "").strip()
        filename = f"Resumen_Inventario_{supplier_code}_{self.date_from}_{self.date_to}.xlsx"

        self.write({
            "archivo": base64.b64encode(xlsx_content),
            "archivo_nombre": filename,
        })

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model={self._name}&id={self.id}&field=archivo&filename_field=archivo_nombre&download=true",
            "target": "self",
        }
