# -*- coding: utf-8 -*-
import base64
import io
import logging
from datetime import datetime

from odoo import api, fields, models, _
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class GdLibroInventarioComparativoWizard(models.TransientModel):
    _name = "gd.libro.inventario.comparativo.wizard"
    _description = "Reporte 2 - Libro de Inventario (Comparativo por proveedor)"

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

    # Rango "Fecha actual"
    date_from_current = fields.Date(string="Desde", required=True)
    date_to_current = fields.Date(string="Hasta", required=True)

    # Rango "Fecha a comparar"
    date_from_compare = fields.Date(string="Desde", required=True)
    date_to_compare = fields.Date(string="Hasta ", required=True)

    archivo = fields.Binary(string="Archivo", readonly=True)
    archivo_nombre = fields.Char(string="Nombre archivo", readonly=True)

    # -------------------------
    # Dominio de proveedores: desde product.supplierinfo
    # -------------------------
    def _get_available_suppliers_domain(self):
        self.ensure_one()
        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("company_id", "in", [False, self.company_id.id]),
            ("partner_id.active", "=", True),
        ])
        partner_ids = supplierinfos.mapped("partner_id").ids
        _logger.info(
            "[GD_R2] Available suppliers from supplierinfo (company=%s): %s",
            self.company_id.id, partner_ids
        )
        return [("id", "in", partner_ids)]

    @api.onchange("company_id")
    def _onchange_company_id(self):
        for w in self:
            return {"domain": {"supplier_id": w._get_available_suppliers_domain()}}

    # -------------------------
    # Helpers read_group
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
    # Validaciones
    # -------------------------
    def _validate_params(self):
        self.ensure_one()
        if not xlsxwriter:
            raise UserError(_("Falta la librería 'xlsxwriter' en tu entorno Python."))

        if self.date_from_current > self.date_to_current:
            raise UserError(_("Rango Actual inválido: 'Desde' no puede ser mayor que 'Hasta'."))

        if self.date_from_compare > self.date_to_compare:
            raise UserError(_("Rango a comparar inválido: 'Desde' no puede ser mayor que 'Hasta'."))

    # -------------------------
    # Proveedor -> Productos (igual que reporte 1)
    # -------------------------
    def _get_product_ids_for_supplier(self):
        self.ensure_one()

        supplierinfos = self.env["product.supplierinfo"].sudo().search([
            ("partner_id", "=", self.supplier_id.id),
            ("company_id", "in", [False, self.company_id.id]),
        ])
        _logger.info("[GD_R2] supplierinfo found: %s", len(supplierinfos))

        product_ids = set()

        # 1) variante
        variant_supplierinfos = supplierinfos.filtered(lambda s: s.product_id)
        for si in variant_supplierinfos:
            product_ids.add(si.product_id.id)
        _logger.info("[GD_R2] supplierinfo variant-level: %s", len(variant_supplierinfos))

        # 2) template => todas variantes
        template_supplierinfos = supplierinfos.filtered(lambda s: not s.product_id)
        tmpl_ids = template_supplierinfos.mapped("product_tmpl_id").ids
        _logger.info("[GD_R2] supplierinfo template-level: %s (templates=%s)", len(template_supplierinfos), tmpl_ids)

        if tmpl_ids:
            variants = self.env["product.product"].sudo().search([
                ("product_tmpl_id", "in", tmpl_ids),
                ("active", "=", True),
            ])
            product_ids.update(variants.ids)

        _logger.info("[GD_R2] FINAL resolved product_ids: %s", list(product_ids))
        return list(product_ids)

    # -------------------------
    # Period stats (neto = out_invoice - out_refund)
    # -------------------------
    def _get_period_stats(self, product_ids, date_from, date_to):
        """Devuelve dict: pid -> {'qty': float, 'total': float}"""
        self.ensure_one()
        if not product_ids:
            return {}

        aml = self.env["account.move.line"].sudo()

        # En tu BD, las líneas reales de producto están como display_type='product'
        base_domain = [
            ("product_id", "in", product_ids),
            ("display_type", "=", "product"),
            ("move_id.company_id", "=", self.company_id.id),
            ("move_id.state", "=", "posted"),
            ("move_id.date", ">=", date_from),   # fecha contable (como en Apuntes contables)
            ("move_id.date", "<=", date_to),
        ]

        inv_domain = base_domain + [("move_id.move_type", "=", "out_invoice")]
        ref_domain = base_domain + [("move_id.move_type", "=", "out_refund")]

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

        _logger.info("[GD_R2] period %s..%s inv_groups=%s ref_groups=%s",
                     date_from, date_to, len(inv_groups), len(ref_groups))

        stats = {}

        # Ventas
        for g in inv_groups:
            pid = g["product_id"][0]
            stats[pid] = {
                "qty": self._rg_sum(g, "quantity"),
                "total": self._rg_sum(g, "price_subtotal"),
            }

        # Devoluciones (neteo)
        for g in ref_groups:
            pid = g["product_id"][0]
            qty = self._rg_sum(g, "quantity")
            total = self._rg_sum(g, "price_subtotal")

            if pid not in stats:
                stats[pid] = {"qty": 0.0, "total": 0.0}

            # Si refund ya viene negativo, sumamos; si viene positivo, restamos.
            if qty < 0 or total < 0:
                stats[pid]["qty"] += qty
                stats[pid]["total"] += total
            else:
                stats[pid]["qty"] -= qty
                stats[pid]["total"] -= total

        # limpia ceros
        stats = {pid: v for pid, v in stats.items()
                 if (abs(v["qty"]) > 1e-9 or abs(v["total"]) > 1e-9)}

        return stats

    # -------------------------
    # Excel (idéntico al Reporte 2)
    # -------------------------
    def _build_xlsx(self, rows):
        """
        rows: lista de dicts:
        {
          product_id,
          qty_current, total_current,
          qty_compare, total_compare
        }
        """
        self.ensure_one()

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {"in_memory": True})
        ws = wb.add_worksheet("Sheet1")  # tu archivo tiene Sheet1

        # Column widths (según tu Excel)
        ws.set_column("A:A", 19.68)
        for col in "BCDEFGHIJ":
            ws.set_column(f"{col}:{col}", 13.0)

        # Row heights (todas 12.8 en tu Excel)
        for r in range(0, 8):
            ws.set_row(r, 12.8)

        # Formats base
        fmt_base = wb.add_format({"font_name": "Arial", "font_size": 10})
        fmt_base_center = wb.add_format({"font_name": "Arial", "font_size": 10, "align": "center"})

        fmt_bold = wb.add_format({"font_name": "Arial", "font_size": 10, "bold": True})

        fmt_time = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "h:mm AM/PM"})

        fmt_header = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "top": 1,      # thin
            "bottom": 6,   # double
            "bg_color": "#DAE3F3",
            "pattern": 1,
            "font_color": "#000000",
        })

        fmt_qty = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "#,##0.00"})
        fmt_total = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "#,##0.00"})

        # Header text
        now_local = fields.Datetime.context_timestamp(self, fields.Datetime.now())
        ws.write(0, 0, "Profit Plus Administrativo", fmt_base)
        ws.write(0, 9, now_local.strftime("%d/%m/%Y"), fmt_base)

        ws.write(1, 0, (self.company_id.name or "").upper(), fmt_base)

        excel_time = datetime(1900, 1, 1, now_local.hour, now_local.minute, 0)
        ws.write_datetime(1, 9, excel_time, fmt_time)

        telefono = self.company_id.phone or ""
        nit = self.company_id.vat or "" 

        ws.write(2, 0, "TEL.:", fmt_base)
        ws.write(2, 1, telefono, fmt_base)
        ws.write(3, 0, "N.I.T..:", fmt_base)
        ws.write(3, 1, nit, fmt_base)

        ws.write(4, 0, "LIBRO DE INVENTARIO", fmt_base)

        proveedor_codigo = self.supplier_id.ref or ""
        ws.write(
            5, 0,
            f"Fecha Actual: {self.date_from_current.strftime('%d/%m/%Y')} Hasta {self.date_to_current.strftime('%d/%m/%Y')} "
            f"vs Fecha anterior {self.date_from_compare.strftime('%d/%m/%Y')} Hasta {self.date_to_compare.strftime('%d/%m/%Y')} "
            f"; Proveedor: {proveedor_codigo}",
            fmt_bold
        )

        # Row 7 labels
        ws.write(6, 3, "Fecha actual", fmt_base_center)      # D7
        ws.write(6, 5, "Fecha a comparar", fmt_base_center)  # F7

        # Row 8 table header
        headers = ["ARTICULO", "MODELO", "DESCRIPCION", "CANTIDAD", "TOTAL", "CANTIDAD", "TOTAL"]
        for col, h in enumerate(headers):
            ws.write(7, col, h, fmt_header)

        # Data rows from row 9 (index 8)
        start_row = 8
        product_ids = [r["product_id"] for r in rows]
        product_map = {p.id: p for p in self.env["product.product"].sudo().browse(product_ids)}

        r = start_row
        for row in rows:
            prod = product_map.get(row["product_id"])
            articulo = (prod.default_code or "") if prod else ""
            modelo = (prod.product_tmpl_id.name or "") if prod and prod.product_tmpl_id else ""
            descripcion = (prod.name or "") if prod else ""

            ws.write(r, 0, articulo, fmt_base)
            ws.write(r, 1, modelo, fmt_base)
            ws.write(r, 2, descripcion, fmt_base)

            ws.write_number(r, 3, float(row["qty_current"]), fmt_qty)
            ws.write_number(r, 4, float(row["total_current"]), fmt_total)

            ws.write_number(r, 5, float(row["qty_compare"]), fmt_qty)
            ws.write_number(r, 6, float(row["total_compare"]), fmt_total)

            r += 1

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
            "[GD_R2] run wizard id=%s company=%s supplier=%s(%s) ranges: curr=%s..%s comp=%s..%s",
            self.id,
            self.company_id.id,
            self.supplier_id.display_name, self.supplier_id.id,
            self.date_from_current, self.date_to_current,
            self.date_from_compare, self.date_to_compare,
        )

        product_ids = self._get_product_ids_for_supplier()
        if not product_ids:
            raise UserError(_("No se encontraron productos vinculados a este proveedor en la pestaña Compras."))

        stats_current = self._get_period_stats(product_ids, self.date_from_current, self.date_to_current)
        stats_compare = self._get_period_stats(product_ids, self.date_from_compare, self.date_to_compare)

        all_pids = set(stats_current.keys()) | set(stats_compare.keys())
        if not all_pids:
            raise UserError(_("No hay movimientos en ninguno de los dos rangos para este proveedor."))

        # Orden: por código de artículo (default_code)
        prods = self.env["product.product"].sudo().browse(list(all_pids))
        product_map = {p.id: p for p in prods}

        def _sort_key(pid):
            p = product_map.get(pid)
            return (p.default_code or "", p.display_name or "")

        sorted_pids = sorted(all_pids, key=_sort_key)

        rows = []
        for pid in sorted_pids:
            c = stats_current.get(pid, {"qty": 0.0, "total": 0.0})
            p = stats_compare.get(pid, {"qty": 0.0, "total": 0.0})
            rows.append({
                "product_id": pid,
                "qty_current": c["qty"],
                "total_current": c["total"],
                "qty_compare": p["qty"],
                "total_compare": p["total"],
            })

        xlsx_content = self._build_xlsx(rows)
        filename = (
            f"LibroInventario_{self.supplier_id.ref or self.supplier_id.id}_"
            f"{self.date_from_current}_{self.date_to_current}_VS_{self.date_from_compare}_{self.date_to_compare}.xlsx"
        )

        self.write({
            "archivo": base64.b64encode(xlsx_content),
            "archivo_nombre": filename,
        })

        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model={self._name}&id={self.id}&field=archivo&filename_field=archivo_nombre&download=true",
            "target": "self",
        }
