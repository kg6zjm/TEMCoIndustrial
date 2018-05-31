# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from datetime import datetime
import lxml.etree as ET
import lxml.builder
from mws import MWSError

from odoo import models, api, _
from odoo.exceptions import UserError
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT


xsi = 'http://www.w3.org/2001/XMLSchema-instance'
E = lxml.builder.ElementMaker(nsmap={'xsi': xsi})
ENCODING = 'iso-8859-1'


class StockPicking(models.Model):
    _inherit = 'stock.picking'

    @api.multi
    def do_transfer(self):
        res = super(StockPicking, self).do_transfer()
        #TODO check if no error when sending the confirmation if order already shipped
        sale_order = self.env['sale.order'].search([('name', '=', self.origin), ('origin', 'like', 'Amazon')])
        if sale_order.product_id.product_tmpl_id.amazon_use:
            Feed = self.env['amazon.instance'].amazon_execute('Feeds')
            feed_xml = self.amazon_order_fulfillment(sale_order)
            try:
                Feed.submit_feed(feed_xml, '_POST_ORDER_FULFILLMENT_DATA_')
            except MWSError as e:
                raise UserError(_('An error occured when updating the shipping status: %s') % e.response)
        return res

    def amazon_order_fulfillment(self, order):
        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('OrderFulfillment'))

        fulfillment_data = E.FulfillmentData(
            E.ShippingMethod('Standard')
        )

        if self.carrier_tracking_ref and self.carrier_id:
            fulfillment_data = E.FulfillmentData(
                E.CarrierName(self.carrier_id.name),
                E.ShipperTrackingNumber(self.carrier_tracking_ref)
            )

        root.append(
            E.Message(
                E.MessageID(str(order.id)),
                E.OrderFulfillment(
                    E.AmazonOrderID(order.client_order_ref),
                    E.FulfillmentDate(
                        datetime.strptime(self.min_date, DEFAULT_SERVER_DATETIME_FORMAT).isoformat('T')
                    ),
                    fulfillment_data
                )
            )
        )
        return ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING)
