# -*- coding: utf-8 -*-

from odoo import models, fields, api

class SaleOrder(models.Model):
    _inherit = 'sale.order'

    def get_payment_methods(self):
        print("self ", self)
        for order in self:
            payments = self.env['payment.method'].search([('active', '=', True)])
        print("payments", payments)
        return payments