# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.
from datetime import datetime

from odoo import models, fields, api


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    access_key = fields.Char(string="Amazon access Key ID")
    secret_key = fields.Char(string="Amazon Secret Key")
    seller_id = fields.Char(string="Amazon Seller ID")
    amazon_site = fields.Many2one("amazon.site", string="Amazon Site Used")
    amazon_currency = fields.Many2one("res.currency", string='Amazon Currency', required=True)
    amazon_sales_team = fields.Many2one("crm.team", string="Sales Team")
    amazon_last_sync = fields.Char("Last Order Synchronization", readonly=True)

    @api.multi
    def set_values(self):
        super(ResConfigSettings, self).set_values()
        set_param = self.env['ir.config_parameter'].sudo().set_param
        set_param('access_key', self.access_key or '')
        set_param('secret_key', self.secret_key or '')
        set_param('seller_id',self.seller_id or '')
        site = self.amazon_site or self.env['amazon.site'].search([], limit=1)
        set_param('amazon_site', site.id)
        set_param('amazon_currency', self.amazon_currency.id)
        amazon_sales_team = self.amazon_sales_team or self.env['crm.team'].search([], limit=1)
        set_param('amazon_sales_team', amazon_sales_team.id)
        # Amazon doesn't handle the microsecond so we strip them off
        amazon_last_sync = self.amazon_last_sync or datetime.now().replace(microsecond=0).isoformat('T')
        set_param('amazon_last_sync', amazon_last_sync)

    @api.model
    def get_values(self):
        res = super(ResConfigSettings, self).get_values()
        get_param = self.env['ir.config_parameter'].sudo().get_param
        res.update(
            access_key = get_param('access_key', default=''),
            secret_key = get_param('secret_key', default=''),
            seller_id = get_param('seller_id', default=''),
            amazon_site = int(get_param('amazon_site', default=self.env['amazon.site'].search([], limit=1))),
            amazon_currency = int(get_param('amazon_currency', default=self.env.ref('base.USD'))),
            amazon_sales_team = int(get_param('amazon_sales_team', default=self.env['crm.team'].search([], limit=1))),
            amazon_last_sync = get_param('amazon_last_sync', default=datetime.now().replace(microsecond=0).isoformat('T'))
        )
        return res

    @api.multi
    def sync_amazon_orders(self):
        return self.env['sale.order'].sync_amazon_orders()

    @api.multi
    def test_amazon_credentials(self):
        return self.env['amazon.instance'].test_amazon_credentials()
