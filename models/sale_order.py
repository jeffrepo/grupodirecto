# -*- coding: utf-8 -*-

from odoo import models, fields, api

class SaleOrder(models.Model):
    _inherit = 'sale.order'

    # def action_open_stock_quant(self):
    #     return {
    #         'name': 'Ubicaciones',
    #         'type': 'ir.actions.act_window',
    #         'res_model': 'stock.quant',
    #         'view_mode': 'list',
    #         'view_id': self.env.ref('stock.view_stock_quant_tree_editable').id,
    #         'target': 'current',
    #         'context': {
    #             'search_default_product': True,
    #             'search_default_location': True
    #         }
    #     }

    def get_discount(self):
        undiscounted_price = 0
        total_discount = 0
        discount = False
        for order in self:
            for line in order.order_line:
                undiscounted_price += line.product_uom_qty * line.price_unit
                total_discount += line.discount
                if discount == False and line.discount:
                    discount = True
        print(f"undiscount_price {undiscounted_price} y total_discount {total_discount}")
        return [discount, undiscounted_price, total_discount]