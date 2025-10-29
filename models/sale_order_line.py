# -*- coding: utf-8 -*-

from odoo import models, fields, api

class SaleOrderLine(models.Model):
    _inherit = 'sale.order.line'

    product_image = fields.Image(
        string="Imagen del Producto",
        max_width=1024,
        max_height=1024,
        help="Imagen asociada al producto en esta línea",
        compute='_compute_product_image',
        store=True
    )
    
    @api.depends('product_template_id')
    def _compute_product_image(self):
        for record in self:
            if record.product_template_id and record.product_template_id.image_1920:
                record.product_image = record.product_template_id.image_1920
            else:
                record.product_image = False