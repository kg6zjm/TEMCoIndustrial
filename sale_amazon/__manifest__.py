# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

{
    'name': "Amazon Connector",
    'summary': "Publish your products on Amazon",
    'description': """
Publish your products on Amazon
==============================================================

The Amazon integrator gives you the opportunity to manage your Odoo's products on Amazon.

Key Features
------------
* Publish products on Amazon
* Revise, relist, end items on Amazon
* Integration with the stock moves
* Automatic creation of sales order and invoices

    """,
    'author': "Odoo SA",
    'website': "https://www.odoo.com",
    'category': 'Sales',
    'version': '1.0',
    'depends': ['base', 'sale_management', 'stock', 'delivery', 'document'],
    'external_dependencies': {'python': ['mws']},
    'data': [
        'security/ir.model.access.csv',
        'data/amazon_data.xml',
        'data/ir_cron_data.xml',
        'wizard/amazon_link_listing_views.xml',
        'views/product_views.xml',
        'views/res_config_views.xml',
    ],
    'application': False,
    'license': 'OEEL-1',
}
