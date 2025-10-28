# -*- coding: utf-8 -*-

from odoo import models, fields

class SaleOrderLine(models.Model):
    _inherit = 'sale.order.line'

    product_image = fields.Image(
        string="Imagen del Producto",
        max_width=1024,
        max_height=1024,
        help="Imagen asociada al producto en esta línea"
    )