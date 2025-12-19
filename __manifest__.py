# -*- coding: utf-8 -*-
{
    'name': "grupodirecto",

    'summary': "Short (1 phrase/line) summary of the module's purpose",

    'description': """
Long description of module's purpose
    """,

    'author': "My Company",
    'website': "https://www.yourcompany.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/15.0/odoo/addons/base/data/ir_module_category_data.xml
    # for the full list
    'category': 'sale',
    'version': '0.1',
    'license': 'LGPL-3',

    # any module necessary for this one to work correctly
    'depends': ['base', 'sale', 'stock', 'purchase','account','product'],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        'views/sale_menus.xml',
        'reports/purchase_order_report_inherit.xml',
        'reports/report_action.xml',
        'reports/sale_order_custom_format_pdf.xml',
        'reports/report_purchasequotation_document_inherit.xml',
        # 'reports/sale_order_template.xml',
        
        'views/product_template_views.xml',
        'views/purchase_order_views.xml',
        "wizards/gd_top_productos_proveedor_views.xml",
        "wizards/gd_libro_inventario_comparativo_views.xml",
        'views/sale_order_views.xml',
        "views/gd_reportes_ventas_menus.xml",
    ],
    # only loaded in demonstration mode
    'demo': [
        'demo/demo.xml',
    ],
}

