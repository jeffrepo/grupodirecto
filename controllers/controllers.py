# -*- coding: utf-8 -*-
# from odoo import http


# class L10nGtGrupodirecto(http.Controller):
#     @http.route('/l10n_gt_grupodirecto/l10n_gt_grupodirecto', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/l10n_gt_grupodirecto/l10n_gt_grupodirecto/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('l10n_gt_grupodirecto.listing', {
#             'root': '/l10n_gt_grupodirecto/l10n_gt_grupodirecto',
#             'objects': http.request.env['l10n_gt_grupodirecto.l10n_gt_grupodirecto'].search([]),
#         })

#     @http.route('/l10n_gt_grupodirecto/l10n_gt_grupodirecto/objects/<model("l10n_gt_grupodirecto.l10n_gt_grupodirecto"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('l10n_gt_grupodirecto.object', {
#             'object': obj
#         })

