# -*- coding: utf-8 -*-

from odoo import models, fields

class PurchaseOrderLine(models.Model):
    _inherit = 'purchase.order.line'

    product_image = fields.Image(
        string="Imagen del Producto",
        max_width=1024,
        max_height=1024,
        help="Imagen asociada al producto en esta línea"
    )