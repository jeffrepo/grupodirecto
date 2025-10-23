# -*- coding: utf-8 -*-

from odoo import models, fields, api

class ProductTemplate(models.Model):
    _inherit = 'product.template'

    x_studio_marca = fields.Char(string="Marca", store=True)

#Eliminar todo estooooooooooooo