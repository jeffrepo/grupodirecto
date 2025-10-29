# -*- coding: utf-8 -*-

from odoo import models, fields, api

class SaleOrder(models.Model):
    _inherit = 'sale.order'

    def action_open_stock_quant(self):
        return {
            'name': 'Ubicaciones',
            'type': 'ir.actions.act_window',
            'res_model': 'stock.quant',
            'view_mode': 'list',
            'view_id': self.env.ref('stock.view_stock_quant_tree_editable').id,
            'target': 'current',
            'context': {
                'search_default_product': True,
                'search_default_location': True
            }
        }