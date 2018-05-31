# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

import lxml.builder

from odoo import models, fields, api, _
from odoo.exceptions import UserError, RedirectWarning

from mws import mws as amazon_mws  # Amazon Marketplace Web Service
xsi = 'http://www.w3.org/2001/XMLSchema-instance'
E = lxml.builder.ElementMaker(nsmap={'xsi': xsi})


# As we use only one instance for now, this model is only
# used for Amazon API methods that could not be put
# in another model.
class AmazonInstance(models.Model):
    _name = "amazon.instance"

    @api.model
    def get_amazon_api(self):

        IrConfigParameter = self.env['ir.config_parameter']
        access_key = IrConfigParameter.get_param('access_key')
        secret_key = IrConfigParameter.get_param('secret_key')
        seller_id = IrConfigParameter.get_param('seller_id')
        amazon_site = IrConfigParameter.get_param('amazon_site')
        if not seller_id or not access_key or not secret_key or not amazon_site:
            action = self.env.ref('sale.action_sale_config_settings')
            raise RedirectWarning(_('One parameter is missing.'),
                                  action.id, _('Configure The Amazon Integrator Now'))
        amazon_site = self.env['amazon.site'].browse(int(amazon_site))
        return {'access_key': access_key, 'secret_key': secret_key,
                'seller_id': seller_id, 'marketplace': amazon_site.name}

    @api.model
    def amazon_execute(self, api_type='MWS', version=False):
        amazon_api = self.get_amazon_api()
        return getattr(amazon_mws, api_type)(amazon_api['access_key'], amazon_api['secret_key'], amazon_api['seller_id'], amazon_api['marketplace'], version=version)

    @api.model
    def generate_root_element(self):
        root = E.AmazonEnvelope(
            E.Header(
                E.DocumentVersion('1.01'),
                E.MerchantIdentifier(self.env['ir.config_parameter'].get_param('seller_id'))
            )
        )
        root.attrib['{%s}noNamespaceSchemaLocation' % (xsi)] = "amzn-envelope.xsd"
        return root

    @api.model
    def test_cron_amazon_credentials(self):
        Sellers = self.amazon_execute('Sellers')
        try:
            response = Sellers.list_marketplace_participations()
            participants = response.parsed.ListParticipations
        except Exception:
            return False
        return True

    @api.model
    def test_amazon_credentials(self):
        Sellers = self.amazon_execute('Sellers')
        try:
            response = Sellers.list_marketplace_participations()
            participants = response.parsed.ListParticipations
        except Exception as e:
            raise UserError(_("Incorrect credentials, we could not find any seller."))
        raise UserError(_("Everything is correctly set up."))

    @api.model
    def get_amazon_marketplaceid(self):
        IrConfigParameter = self.env['ir.config_parameter']
        amazon_site = self.env['amazon.site'].browse(int(IrConfigParameter.get_param('amazon_site')))
        return amazon_site.amazon_id

    @api.model
    def report_by_list_of_dict(self, report):

        # convert bytecode to utf8 and ignore non-ascii characters.
        report = report.decode("utf-8", 'ignore')
        results = []
        lines = report.split('\n')

        # First line is field names
        fields = lines[0].split('\t')
        for line in lines[1:-1]:
            line_dict = {}
            values = line.split('\t')
            for i in range(len(values)):
                line_dict[fields[i]] = values[i]

            results.append(line_dict)

        return results


class amazon_site(models.Model):
    _name = "amazon.site"

    name = fields.Char("Name", readonly=True)
    amazon_id = fields.Char("Amazon Marketplace ID", readonly=True)
