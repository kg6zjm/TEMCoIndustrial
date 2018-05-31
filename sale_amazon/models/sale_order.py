# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from datetime import datetime
import lxml.builder
import lxml.etree as ET
from mws import MWSError

from odoo import models, fields, api, _
from odoo.exceptions import UserError

xsi = 'http://www.w3.org/2001/XMLSchema-instance'
E = lxml.builder.ElementMaker(nsmap={'xsi': xsi})
ENCODING = 'iso-8859-1'


class SaleOrder(models.Model):
    _inherit = "sale.order"

    # Override field to make it copy=True
    client_order_ref = fields.Char(string='Customer Reference', copy=True)

    @api.multi
    def action_cancel(self):
        res = super(SaleOrder, self).action_cancel()
        self._update_order_status('Failure')
        return res

    @api.multi
    def action_confirm(self):
        res = super(SaleOrder, self).action_confirm()
        self._update_order_status('Success')
        return res

    def _update_order_status(self, status_code=None):
        orders = self.filtered(lambda so: so.origin and 'amazon' in so.origin)
        if not orders:
            return False
        Feed = self.env['amazon.instance'].amazon_execute('Feeds')
        feed_xml = orders._order_acknowledgement(status_code)
        try:
            Feed.submit_feed(feed_xml, '_POST_ORDER_ACKNOWLEDGEMENT_DATA_')
        except MWSError as e:
            raise UserError(_('An error occured when updating the sale order status: %s') % e.response)

    def _order_acknowledgement(self, status_code=None):
        "when order is confirm or canceled acknowledgment sent to amazon"
        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('OrderAcknowledgement'))
        for order in self:
            if order.client_order_ref:
                root.append(
                    E.Message(
                        E.MessageID(str(order.id)),
                        E.OrderAcknowledgement(
                            E.AmazonOrderID(order.client_order_ref),
                            E.StatusCode(status_code)
                        )
                    )
                )
        return ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING)

    @api.model
    def sync_amazon_orders(self):
        SaleOrderLine = self.env['sale.order.line']
        Order = self.env['amazon.instance'].amazon_execute('Orders', '2013-09-01')
        for amazon_order in self._list_orders(Order):
            amazon_order_id = amazon_order['AmazonOrderId']['value']
            if not self.env['sale.order'].search([('client_order_ref', '=', amazon_order_id)]):
                order_items = self._list_order_items(Order, amazon_order_id)
                order_items_sku = [i['SellerSKU']['value'] for i in order_items]
                if self.env['product.product'].search([('amazon_sku', 'in', order_items_sku)]):
                    order = self._create_sale_order(amazon_order, amazon_order_id)
                    currency = self.env['res.currency'].search([('name', '=', amazon_order['OrderTotal']['CurrencyCode']['value'])])
                    if self.env['ir.config_parameter'].get_param('amazon_sales_team'):
                        order.team_id = int(self.env['ir.config_parameter'].get_param('amazon_sales_team'))
                    for order_item in self._list_order_items(Order, amazon_order_id):
                        variant = self.env['product.product'].search([
                            ('amazon_sku', '=', order_item['SellerSKU']['value']),
                            ('amazon_use', '=', True)
                        ])
                        if len(variant) > 1:
                            raise UserError(_("There are products with the same Amazon SKU. Please avoid this before synchronizing the orders."))
                        if variant:
                            company_id = self.env.user.company_id
                            IrDefault = self.env['ir.default']
                            if variant.taxes_id:
                                taxes_id = variant.taxes_id.mapped('id')
                            else:
                                taxes_id = IrDefault.get('product.template', 'taxes_id', company_id=company_id.id)
                            sol = SaleOrderLine.create({
                                'product_id': variant.id,
                                'order_id': order.id,
                                'product_uom_qty': float(order_item['QuantityOrdered']['value']),
                                'price_unit': currency.compute(
                                    (float(order_item['ItemPrice']['Amount']['value'])/float(order_item['QuantityOrdered']['value'])),
                                    self.env.user.company_id.currency_id),
                                'discount': (float(order_item['PromotionDiscount']['Amount']['value'])*100)/float(order_item['ItemPrice']['Amount']['value']),
                                'tax_id': [(6, 0, taxes_id)] if taxes_id else False,
                            })
                            # Since Amazon only return the amount of the tax and not the percentage
                            # Use the field fixed_sale_tax to stock this amount
                            # The field is used later in the compute_all method of the account.tax
                            if order_item.get('ItemTax'):
                                tax_amount = float(order_item['ItemTax']['Amount']['value'])
                                if tax_amount > 0:
                                    amazon_tax_id = self.env.ref('sale_amazon.amazon_sale_tax').id
                                    sol.write({
                                        'fixed_sale_tax': tax_amount,
                                        'tax_id': [(4, amazon_tax_id, None)]
                                    })
                            if order_item.get('ShippingPrice'):
                                shipping_name = amazon_order['ShipServiceLevel']['value']
                                shipping_product = self.env['product.template'].search([('name', '=', shipping_name)])
                                if not shipping_product:
                                    shipping_product = self.env['product.template'].create({
                                        'name': shipping_name,
                                        'uom_id': self.env.ref('product.product_uom_unit').id,
                                        'type': 'service',
                                        'categ_id': self.env.ref('sale_amazon.product_category_amazon').id,
                                    })
                                if shipping_product.taxes_id:
                                    taxes_id = shipping_product.taxes_id.mapped('id')
                                else:
                                    taxes_id = IrDefault.get('product.template', 'taxes_id', company_id=company_id.id)
                                discount = (float(order_item['ShippingDiscount']['Amount']['value'])*100)\
                                    / float(order_item['ShippingPrice']['Amount']['value'])
                                shipping_sol = SaleOrderLine.create({
                                    'order_id': order.id,
                                    'name': shipping_name,
                                    'product_id': shipping_product.product_variant_ids[0].id,
                                    'product_uom': self.env.ref('product.product_uom_unit').id,
                                    'product_uom_qty': 1,
                                    'price_unit': currency.compute(
                                        float(order_item['ShippingPrice']['Amount']['value']),
                                        self.env.user.company_id.currency_id),
                                    'is_delivery': True,
                                    'discount': discount,
                                    'tax_id': [(6, 0, taxes_id)] if taxes_id else False,
                                })
                                if order_item.get('ShippingTax'):
                                    tax_amount = float(order_item['ShippingTax']['Amount']['value'])
                                    if tax_amount > 0:
                                        amazon_tax_id = self.env.ref('sale_amazon.amazon_sale_tax').id
                                        shipping_sol.write({
                                            'fixed_sale_tax': tax_amount,
                                            'tax_id': [(4, amazon_tax_id, None)]
                                        })
                                variant_data = {
                                    'amazon_quantity_sold': variant.amazon_quantity_sold + int(order_item['QuantityOrdered']['value']),
                                    'amazon_last_sync': datetime.now(),
                                }
                                if not variant.product_tmpl_id.amazon_sync_stock:
                                    variant_data.update(amazon_quantity=variant.amazon_quantity - int(order_item['QuantityOrdered']['value']))
                                variant.write(variant_data)

                    order.action_confirm()
                    if amazon_order['OrderStatus']['value'] == 'Shipped':
                        order.picking_ids.do_transfer()
                    order.action_invoice_create(final=True)
        self.env['ir.config_parameter'].set_param('amazon_last_sync', datetime.now().isoformat('T'))

    def _create_sale_order(self, order_dict, amazon_order_id):
        partner = self.env['res.partner'].search([
            ('email', '=', order_dict['BuyerEmail']['value']),
            ('type', '=', 'contact'),
        ])
        if not partner:
            create_data = {
                'name': order_dict['BuyerName']['value'],
                'email': order_dict['BuyerEmail']['value'],
                'ref': 'Amazon',
            }
            partner = self.env['res.partner'].create(create_data)
        if 'ShippingAddress' in order_dict:
            address = order_dict['ShippingAddress']
            partner_data = {
                'street': address['AddressLine1']['value'],
                'city': address['City']['value'],
                'zip': address['PostalCode']['value'],
                'country_id': self.env['res.country'].search([
                    ('code', '=', address['CountryCode']['value'])
                ], limit=1).id,
                'state_id': self.env['res.country.state'].search([
                    ('code', '=', address['StateOrRegion']['value'])
                ], limit=1).id,
            }
            shipping_partner = self.env['res.partner'].search([
                ('name', '=', address['Name']['value']),
                ('email', '=', order_dict['BuyerEmail']['value']),
                ('street', '=', partner_data['street']),
                ('city', '=', partner_data['city']),
                ('zip', '=', partner_data['zip']),
                ('type', '=', 'delivery'),
            ])
            partner.write(partner_data)
            if not shipping_partner:
                partner_data['parent_id'] = partner.id
                partner_data['name'] = address['Name']['value']
                partner_data['email'] = order_dict['BuyerEmail']['value']
                partner_data['ref'] = 'Amazon'
                partner_data['type'] = 'delivery'
                shipping_partner = self.env['res.partner'].create(partner_data)
        fp_id = self.env['account.fiscal.position'].get_fiscal_position(partner.id)
        if fp_id:
            partner.property_account_position_id = fp_id
        return self.env['sale.order'].create({
            'partner_id': partner.id,
            'state': 'draft',
            'client_order_ref': amazon_order_id,
            'origin': 'Amazon ' + amazon_order_id,
            'fiscal_position_id': fp_id if fp_id else False,
            'partner_shipping_id': shipping_partner.id,
        })

    def _list_orders(self, Order):
        """ Get list of orders.
            Amazon returns a limited number of order by response.
            If `NextToken` is returned as well, it means that we
            need to fetch again the list of orders to get all the orders.
        """
        last_sync = self.env['ir.config_parameter'].get_param('amazon_last_sync')
        try:
            orders_list = Order.list_orders([self.env['amazon.instance'].get_amazon_marketplaceid()], created_after=last_sync, orderstatus=('Unshipped', 'PartiallyShipped', 'Shipped')).parsed
        except MWSError as e:
            raise UserError(_('An error occured during the synchronization: %s') % e.response.text)
        if not orders_list['Orders']:
            return []
        amazon_orders_list = orders_list['Orders']['Order']
        if not isinstance(amazon_orders_list, list):
            amazon_orders_list = [amazon_orders_list]

        def get_list_by_next_token(token):
            # Recursively get order list until there are no more orders to fetch
            order_token_list = Order.list_orders_by_next_token(token).parsed
            orders_token = orders_list['Orders']['Order']
            if not isinstance(orders_token, list):
                orders_token = [orders_token]
            if order_token_list.get('NextToken'):
                orders_token += get_list_by_next_token(order_token_list['NextToken']['value'])
            return orders_token

        # if `NextToken` is returned, there are more orders to fetch
        if orders_list.get('NextToken'):
            amazon_orders_list += get_list_by_next_token(orders_list['NextToken']['value'])
        return amazon_orders_list

    def _list_order_items(self, Order, amazon_order_id):
        """ Get order items.
            Amazon returns a limited number of order by response.
            If `NextToken` is returned as well, it means that we
            need to fetch again the list of order items to get all the items.
        """
        try:
            order_items = Order.list_order_items(amazon_order_id).parsed
        except MWSError as e:
            raise UserError(_('An error occured during the synchronization: %s') % e.response)
        item_lists = order_items['OrderItems']['OrderItem']
        if not isinstance(item_lists, list):
            item_lists = [item_lists]

        def get_items_by_next_token(token):
            # Recursively get items list until there are no more items to fetch
            items_token_list = Order.list_order_items_by_next_token(token).parsed
            items_token = items_token_list['OrderItems']['OrderItem']
            if not isinstance(items_token, list):
                items_token = [items_token]
            if items_token_list.get('NextToken'):
                items_token += get_items_by_next_token(items_token_list['NextToken']['value'])
            return items_token
        if order_items.get('NextToken'):
            item_lists += get_items_by_next_token(order_items['NextToken']['value'])
        return item_lists


class SaleOrderLine(models.Model):
    _inherit = "sale.order.line"

    fixed_sale_tax = fields.Float('Fixed Sale Tax Amount')

    @api.depends('product_uom_qty', 'discount', 'price_unit', 'tax_id')
    def _compute_amount(self):
        """
        Compute the amounts of the SO line.
        """
        for line in self:
            price = line.price_unit * (1 - (line.discount or 0.0) / 100.0)
            taxes = line.tax_id.compute_all(
                price,
                line.order_id.currency_id,
                line.product_uom_qty,
                product=line.product_id,
                partner=line.order_id.partner_id,
                set_amount=line.fixed_sale_tax)
            line.update({
                'price_tax': taxes['total_included'] - taxes['total_excluded'],
                'price_total': taxes['total_included'],
                'price_subtotal': taxes['total_excluded'],
            })


class AccountAmazonSalesTax(models.Model):
    _inherit = "account.tax"

    amount_type = fields.Selection(selection_add=[('fixed_set', 'Fixed amount set on model')])

    @api.multi
    def compute_all(self, price_unit, currency=None, quantity=1.0, product=None, partner=None, set_amount=0.0):
        taxes = self.filtered(lambda t: t.amount_type != 'fixed_set')
        result = super(AccountAmazonSalesTax, taxes).compute_all(price_unit, currency, quantity, product, partner)
        for tax in self.filtered(lambda t: t.amount_type == 'fixed_set'):
            result['total_included'] += set_amount
            if not currency:
                currency = tax.company_id.currency_id
            base = round(price_unit * quantity, currency.decimal_places)
            result['taxes'].append({
                'id': tax.id,
                'name': tax.with_context(**{'lang': partner.lang} if partner else {}).name,
                'amount': set_amount,
                'sequence': tax.sequence,
                'account_id': tax.account_id.id,
                'refund_account_id': tax.refund_account_id.id,
                'analytic': tax.analytic,
                'base': base,
            })
        return result

class AccountAmazonInvoice(models.Model):
    _inherit = "account.invoice"

    @api.multi
    def get_taxes_values(self):
        tax_grouped = {}
        for line in self.invoice_line_ids:
            price_unit = line.price_unit * (1 - (line.discount or 0.0) / 100.0)
            set_amount = sum(line.sale_line_ids.mapped('fixed_sale_tax'))
            taxes = line.invoice_line_tax_ids.compute_all(price_unit, self.currency_id, line.quantity, line.product_id, self.partner_id, set_amount=set_amount)['taxes']
            for tax in taxes:
                val = self._prepare_tax_line_vals(line, tax)
                key = self.env['account.tax'].browse(tax['id']).get_grouping_key(val)

                if key not in tax_grouped:
                    tax_grouped[key] = val
                else:
                    tax_grouped[key]['amount'] += val['amount']
                    tax_grouped[key]['base'] += val['base']
        return tax_grouped
