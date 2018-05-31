# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

import lxml.builder
import lxml.etree as ET
from mws import MWSError

from odoo import models, fields, api, _
from odoo.exceptions import UserError

xsi = 'http://www.w3.org/2001/XMLSchema-instance'
E = lxml.builder.ElementMaker(nsmap={'xsi': xsi})
ENCODING = 'iso-8859-1'


class ProductTemplate(models.Model):
    _inherit = "product.template"

    amazon_use = fields.Boolean('Use Amazon', related='product_variant_ids.amazon_use')
    amazon_state = fields.Selection([
        ('not_published', 'Not Published'),
        ('to_publish', 'Waiting to be Published'),
        ('published', 'Published'),
        ('to_revise', 'Waiting to be Revised'),
        ('quantity_missing', 'Quantity Missing'),  # used with link with existing if FBM because inventory is async.
        ('missing_info', 'Missing Info'),
        ('listing_ended', 'Listing Ended'),
        ('to_delete', 'Waiting to be Deleted'),
        ('wait_feed_result', 'Waiting for Feed Submission Result'),
        ('wait_request_report', 'Waiting for Request Report Result'),
    ], string='Amazon State', default='not_published', compute='_compute_amazon_state', readonly=True, store=True)
    amazon_product_type = fields.Selection([
        ('ASIN', 'ASIN'), ('ISBN', 'ISBN'), ('UPC', 'UPC'), ('EAN', 'EAN')],
        string='Amazon Type', related='product_variant_ids.amazon_product_type', store=True)
    amazon_product_type_value = fields.Char('Amazon Type Value', related='product_variant_ids.amazon_product_type_value', store=True)
    amazon_sku = fields.Char('Amazon SKU', related='product_variant_ids.amazon_sku', store=True, help="Unique Identifier. If you don't enter a SKU Amazon will create one for you")
    amazon_condition = fields.Selection([
        ('New', 'New'), ('UsedLikeNew', 'Used Like New'), ('UsedVeryGood', 'Used Very Good'), ('UsedGood', 'Used Good'),
        ('UsedAcceptable', 'Used Acceptable'), ('CollectibleLikeNew', 'Collectible Like New'), ('CollectibleVeryGood', 'Collectible Very Good'),
        ('CollectibleGood', 'Collectible Good'), ('CollectibleAcceptable', 'Collectible Acceptable'), ('Refurbished', 'Refurbished'), ('Club', 'Club')
    ], string="Amazon Condition", related='product_variant_ids.amazon_condition', store=True)
    amazon_fixed_price = fields.Float('Amazon Fixed Price', related='product_variant_ids.amazon_fixed_price', store=True)
    amazon_shipping_price = fields.Float('Amazon Shipping Price')
    amazon_quantity = fields.Integer(string='Amazon Quantity', related='product_variant_ids.amazon_quantity', store=True)
    amazon_quantity_sold = fields.Integer(related='product_variant_ids.amazon_quantity_sold', store=True)
    amazon_product_link = fields.Char('Product Link', compute='_compute_amazon_product_link', store=True, readonly=True)
    amazon_fulfillment = fields.Selection([
        ('MFN', 'Fulfilled by Merchant'), ('AFN', 'Fulfilled by Amazon')
    ], string='Amazon Fulfillment', default='MFN')
    amazon_fulfillment_center_id = fields.Selection([
        ('DEFAULT', 'DEFAULT'), ('AMAZON_NA', 'AMAZON_NA'), ('AMAZON_EU', 'AMAZON_EU'), ('AMAZON_JP', 'AMAZON_JP'), ('AMAZON_CN', 'AMAZON_CN'), ('AMAZON_IN', 'AMAZON_IN')
    ], string='Amazon Fulfillment Center ID', default='DEFAULT')
    amazon_sync_stock = fields.Boolean(string="Use The Stock's Quantity", default=False)
    amazon_last_sync = fields.Datetime('Last Sync Date')

    @api.one
    @api.depends('amazon_product_type_value', 'amazon_product_type')
    def _compute_amazon_product_link(self):
        if self.amazon_product_type and self.amazon_product_type_value:
            amazon_site_id = self.env['ir.config_parameter'].get_param('amazon_site')
            amazon_extension = str.lower(str(self.env['amazon.site'].browse(int(amazon_site_id)).name))
            if amazon_extension == 'us':
                amazon_extension = 'com'
            elif amazon_extension == 'uk':
                amazon_extension = 'co.uk'
            amazon_product_link = 'http://www.amazon.' + amazon_extension
            if self.amazon_product_type == 'ASIN':
                self.amazon_product_link = amazon_product_link + '/dp/' + self.amazon_product_type_value
            else:
                self.amazon_product_link = amazon_product_link + '/s?url=search-alias%3Daps&field-keywords=' + self.amazon_product_type_value

    @api.one
    @api.depends('product_variant_ids.amazon_state')
    def _compute_amazon_state(self):
        if self.product_variant_count == 1:
            self.amazon_state = self.product_variant_ids.amazon_state
        else:
            variant_states = [v.amazon_state for v in self.product_variant_ids]
            if 'quantity_missing' in variant_states:
                self.amazon_state = 'quantity_missing'
            elif 'missing_info' in variant_states:
                self.amazon_state = 'missing_info'
            elif 'listing_ended' in variant_states:
                self.amazon_state = 'listing_ended'
            elif 'to_publish' in variant_states:
                self.amazon_state = 'to_publish'
            elif 'to_revise' in variant_states:
                self.amazon_state = 'to_revise'
            elif 'to_delete' in variant_states:
                self.amazon_state = 'to_delete'
            elif 'wait_feed_result' in variant_states:
                self.amazon_state = 'wait_feed_result'
            elif 'wait_request_report' in variant_states:
                self.amazon_state = 'wait_request_report'
            elif 'published' in variant_states:
                self.amazon_state = 'published'
            else:
                self.amazon_state = 'not_published'

    @api.multi
    def publish_product_amazon_to_queue(self):
        # The products are pushed asynchronously so we just add them in a queue
        self.ensure_one()
        for v in self.product_variant_ids:
            if v.amazon_state != 'published':
                v.amazon_state = 'to_publish'

    @api.multi
    def revise_product_amazon_to_queue(self):
        self.ensure_one()
        for v in self.product_variant_ids:
            if v.amazon_state != 'not_published':
                v.amazon_state = 'to_revise'

    @api.multi
    def end_listing_product_amazon(self):
        # TODO: improve
        # TODO: add state waiting_ending
        # the result looks like :
        #  '_response_dict': {'ResponseMetadata': {'RequestId': {'value': 'fa208ade-4ed4-4273-b59f-c3a117db387b'}},
        # 'SubmitFeedResult': {'FeedSubmissionInfo': {'FeedProcessingStatus': {'value': '_SUBMITTED_'},
        #                                             'FeedSubmissionId': {'value': '51099016885'},
        #                                             'FeedType': {'value': '_POST_PRODUCT_DATA_'},
        #                                             'SubmittedDate': {'value': '2016-03-25T08:15:19+00:00'}}}}
        # How to check if the request is really accepted?
        self.ensure_one()
        Feed = self.env['amazon.instance'].amazon_execute('Feeds')
        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('Product'))
        variants = self.product_variant_ids.filtered(lambda x: x.amazon_state != 'not_published')
        for variant in variants:
            product_xml = E.Product()
            if variant.amazon_sku:
                product_xml.append(E.SKU(variant.amazon_sku))
            root.append(
                E.Message(
                    E.MessageID(str(variant.id)),
                    E.OperationType('Delete'),
                    product_xml
                )
            )
        try:
            response = Feed.submit_feed(
                ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING),
                '_POST_PRODUCT_DATA_'
            )
        except MWSError as e:
            raise UserError(_('An error occured when ending the listing: %s') % e.response)
        variants._set_waiting_product(response.parsed.FeedSubmissionInfo.FeedSubmissionId, 'delete')
        return response.parsed.FeedSubmissionInfo.FeedSubmissionId

    def _find_corresponding_products(self):

        AmazonProducts = self.env['amazon.instance'].amazon_execute('Products')
        marketplace_id = self.env['amazon.instance'].get_amazon_marketplaceid()
        try:
            response = AmazonProducts.get_matching_product_for_id(
                marketplace_id, self.amazon_product_type, [self.amazon_product_type_value])
        except MWSError as e:
            raise UserError(_('An error occured when linking the listing: %s') % e.response)
        if 'Products' in response.parsed:
            return response.parsed.Products
        else:
            raise UserError(response.parsed.Error.Message)

    @api.multi
    def unlink_listing_product_amazon(self):
        for product in self:
            if product.amazon_use:
                vals = {
                    'amazon_state': 'not_published',
                    'amazon_use': False,
                    'amazon_product_type_value': False,
                    'amazon_sku': False,
                    'amazon_product_link': False,
                }
                product.write(vals)
                product.product_variant_ids.write(vals)


class ProductProduct(models.Model):
    _inherit = "product.product"

    amazon_use = fields.Boolean('Use Amazon')
    amazon_state = fields.Selection([
        ('not_published', 'Not Published'),
        ('to_publish', 'Waiting to be Published'),
        ('published', 'Published'),
        ('to_revise', 'Waiting to be Revised'),
        ('quantity_missing', 'Quantity Missing'),  # used with link with existing if FBM because inventory is async.
        ('missing_info', 'Missing Info'),
        ('listing_ended', 'Listing Ended'),
        ('to_delete', 'Waiting to be Deleted'),
        ('wait_feed_result', 'Waiting for Feed Submission Result'),
        ('wait_request_report', 'Waiting for Request Report Result'),
    ], string='Amazon State', default='not_published', readonly=True)
    amazon_product_type = fields.Selection([('ASIN', 'ASIN'), ('ISBN', 'ISBN'), ('UPC', 'UPC'), ('EAN', 'EAN')], string='Type')
    amazon_product_type_value = fields.Char('Amazon Type Value')
    amazon_sku = fields.Char('Amazon SKU', help="Unique Identifier. If you don't enter a SKU Amazon will create one for you")
    amazon_quantity = fields.Integer(string='Amazon Quantity', default=0)
    amazon_quantity_sold = fields.Integer('Amazon Quantity Sold', readonly=True)
    amazon_fixed_price = fields.Float('Amazon Fixed Price')
    amazon_shipping_price = fields.Float('Amazon Shipping Price')
    amazon_condition = fields.Selection([
        ('New', 'New'), ('UsedLikeNew', 'Used Like New'), ('UsedVeryGood', 'Used Very Good'), ('UsedGood', 'UsedGood'),
        ('UsedAcceptable', 'Used Acceptable'), ('CollectibleLikeNew', 'Collectible Like New'), ('CollectibleVeryGood', 'Collectible Very Good'),
        ('CollectibleGood', 'Collectible Good'), ('CollectibleAcceptable', 'Collectible Acceptable'), ('Refurbished', 'Refurbished'), ('Club', 'Club')
    ])
    amazon_feed_id = fields.Char('Amazon Submission Feed ID')
    amazon_report_id = fields.Char('Amazon Report Request ID')
    amazon_waiting_feed_type = fields.Char('Amazon Waiting Feed Type')
    amazon_report_type = fields.Char('Amazon Report Request Type')

    @api.model
    def amazon_cron(self):
        if not self.env['amazon.instance'].test_cron_amazon_credentials():
            return
        products_to_revise = self.search([('amazon_state', '=', 'to_revise')])
        print('Products to revise : %s' % products_to_revise)
        if products_to_revise:
            products_to_revise.revise_products_amazon()
        products_to_publish = self.search([('amazon_state', '=', 'to_publish')])
        print('Variants to publish : %s' % products_to_publish)
        if products_to_publish:
            products_to_publish.publish_products_amazon()
        product_to_retrieve = self.search([('amazon_state', '=', 'quantity_missing')])
        print('Products to retrieve quantity : %s' % product_to_retrieve)
        for product in product_to_retrieve:
            product._list_inventory_supply_by_merchant()
        waiting_feed_product = self.search([('amazon_state', '=', 'wait_feed_result')])
        print('Variants waiting feed: %s' % waiting_feed_product)
        Feed = self.env['amazon.instance'].amazon_execute('Feeds')
        for product in waiting_feed_product:
            product._process_feed_submission_result(Feed,product.amazon_feed_id, product.amazon_waiting_feed_type)
        waiting_report_product = self.search([('amazon_state', '=', 'wait_request_report')])
        print('Variants waiting report: %s' % waiting_report_product)
        Reports = self.env['amazon.instance'].amazon_execute('Reports')
        for product in waiting_report_product:
            product._process_report_request(Reports, product.amazon_report_id, product.amazon_report_type)

    def publish_products_amazon(self):

        Feed = self.env['amazon.instance'].amazon_execute('Feeds')

        product_feed_submission_id = self._submit_product_feed(Feed)
        listed_products = self._process_feed_submission_result(Feed, product_feed_submission_id, 'product')
        print('Listed product : %s' % listed_products)
        self._cr.commit()

    def revise_products_amazon(self):

        Feed = self.env['amazon.instance'].amazon_execute('Feeds')
        inventory_feed_submission_id = self._submit_inventory_feed(Feed)
        self._process_feed_submission_result(Feed, inventory_feed_submission_id, 'inventory')

    def _list_inventory_supply_by_merchant(self):

        Reports = self.env['amazon.instance'].amazon_execute('Reports')
        marketplaceid = self.env['amazon.instance'].get_amazon_marketplaceid()

        report_type = '_GET_MERCHANT_LISTINGS_DATA_'
        try:
            report_request = Reports.request_report(report_type, marketplaceids=(marketplaceid,))
        except MWSError as e:
            raise UserError(_('An error occured when requesting the report: %s') % e.response)
        report_request_id = report_request.parsed['ReportRequestInfo']['ReportRequestId']['value']
        self._process_report_request(Reports, report_request_id, report_type)

    def _process_report_request(self, Reports, report_request_id, report_type):
        try:
            report_request_list = Reports.get_report_request_list([report_request_id])
            status = report_request_list.parsed.ReportRequestInfo.ReportProcessingStatus
            if status == '_DONE_':
                generated_report_id = report_request_list.parsed.ReportRequestInfo.GeneratedReportId
                if not generated_report_id:
                    report_list = Reports.get_report_list([report_request_id], types=[report_type])
                    generated_report_id = report_list.parsed.ReportInfo.ReportId
                result = Reports.get_report(generated_report_id)
                response_msg = result.parsed
                inventory = self.env['amazon.instance'].report_by_list_of_dict(response_msg)
                variant_found = False
                for line in inventory:
                    sku = line.get('seller-sku')
                    variant = self.search([
                        ('amazon_sku', '=', sku),
                    ])
                    # We check if the variant found is in the record set
                    # 'product_to_retrieve' or 'waiting_report_product'
                    if variant and variant in self:
                        variant.write({
                            'amazon_quantity': int(line.get('quantity')),
                            'amazon_state': 'published'
                        })
                        variant_found = True
                if not variant_found:
                    self.product_tmpl_id.message_post(
                        _("Impossible to retrieve the Amazon Quantity for Variant %s. The product was not found in your inventory." % self.display_name))
                    self.write({
                        'amazon_state': 'quantity_missing',
                    })
                return response_msg
            elif status == '_CANCELLED_':
                self.write({
                    'amazon_state': 'quantity_missing',
                })
            elif status == '_DONE_NO_DATA_':
                self.write({
                    'amazon_state': 'missing_info'
                })
                self.product_tmpl_id.message_post(_("Impossible to retrieve the Amazon quantity for Variant %s." % self.display_name))
            elif status in ['_IN_PROGRESS_', '_SUBMITTED_']:
                self.write({
                    'amazon_report_id': report_request_id,
                    'amazon_report_type': report_type,
                    'amazon_state': 'wait_request_report',
                })
        except MWSError as e:
            raise UserError(_('An error occured when processing the report request: %s') % e.response)

    def _list_inventory_supply_by_amazon(self):

        AmazonInventory = self.env['amazon.instance'].amazon_execute('Inventory')
        # TODO: do we need "Details" ? --> FNSKU, date, etc.
        try:
            response = AmazonInventory.list_inventory_supply(skus=[self.amazon_sku])
        except MWSError as e:
            raise UserError(_('An error occured when listing the inventory: %s') % e.response)

        if 'InventorySupplyList' in response.parsed:
            return response.parsed.InventorySupplyList
        else:
            raise UserError(response.parsed.Error.Message)

    def _submit_product_feed(self, Feed):
        amazon_site = self.env['amazon.site'].browse(int(self.env['ir.config_parameter'].get_param('amazon_site')))

        feed_xml = self._list_new_products()
        try:
            response = Feed.submit_feed(feed_xml, '_POST_PRODUCT_DATA_', (amazon_site.amazon_id,))
        except MWSError as e:
            raise UserError(_('An error occured when submitting the product feed: %s') % e.response)

        if 'FeedSubmissionInfo' in response.parsed:
            return response.parsed.FeedSubmissionInfo.FeedSubmissionId
        else:
            raise UserError(response.parsed.Error.Message)

    def _set_waiting_product(self, feed_submission_id, feed_type):
        self.write({
            'amazon_feed_id': feed_submission_id,
            'amazon_waiting_feed_type': feed_type,
            'amazon_state': 'wait_feed_result',
        })

    def _process_feed_submission_result(self, Feed, feed_submission_id, feed_type):
        try:
            response = Feed.get_feed_submission_result(feed_submission_id)
            response_msg = response._response_dict.Message
            from pprint import pprint
            pprint(response_msg)
        except MWSError as e:
            root = ET.fromstring(e.response.text)
            if root.find('.//{http://mws.amazonaws.com/doc/2009-01-01/}Code').text == 'FeedProcessingResultNotReady':
                self._set_waiting_product(feed_submission_id, feed_type)
                return
            raise UserError(_('An error occured when processing the feed subsmission result: %s') % e.response)
        except Exception as e:
            self._set_waiting_product(feed_submission_id, feed_type)
            return

        published_products = self

        try:
            report_summary = response_msg.ProcessingReport.ProcessingSummary
            if report_summary.MessagesProcessed != report_summary.MessagesSuccessful:
                # There is a least one problem
                response_results = response_msg.ProcessingReport.Result
                response_results = isinstance(response_results, (list)) and response_results or [response_results]
                for result in response_results:
                    variant = self.browse(int(result.MessageID))
                    result_code = result.ResultCode

                    if result_code == 'Error':
                        # The item hasn't been published correctly, we only return published products
                        published_products -= variant
                        error_message = result.ResultDescription
                        variant.product_tmpl_id.message_post(
                            _("The %s could not be listed on Amazon. You should fix the following error and try to push again : %s") % (feed_type, error_message)
                        )
                        if feed_type == 'product':
                            variant.amazon_state = 'not_published'
                        else:
                            variant.amazon_state = 'missing_info'
        except Exception as e:
            self._set_waiting_product(feed_submission_id, feed_type)
            return

        if len(published_products):
            if feed_type == 'delete':
                print('Deleted products: %s' % published_products)
                published_products.write({'amazon_state': 'not_published'})
                for variant in published_products:
                    variant.product_tmpl_id.message_post(_("The %s Product has been removed from your Amazon inventory." % variant.display_name))
            elif feed_type == 'product':
                print('Start inventory for product : %s' % published_products)
                inventory_feed_submission_id = published_products._submit_inventory_feed(Feed)
                self._process_feed_submission_result(Feed, inventory_feed_submission_id, 'inventory')
                for variant in published_products:
                    variant.product_tmpl_id.message_post(_("The %s product is now listed on your Amazon inventory." % variant.display_name))
            elif feed_type == 'inventory':
                print('Start price for product : %s' % published_products)
                price_feed_submission_id = published_products._submit_price_feed(Feed)
                self._process_feed_submission_result(Feed, price_feed_submission_id, 'price')
            else:
                print('Publish finish for product : %s' % published_products)
                published_products.write({'amazon_state': 'published'})
        return published_products

    def _submit_inventory_feed(self, Feed):
        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('Inventory'))
        for v in self:
            p = v.product_tmpl_id
            if p.amazon_sync_stock:
                v.amazon_quantity = max(int(v.virtual_available), 0)
            root.append(
                E.Message(
                    E.MessageID(str(v.id)),
                    E.OperationType('Update'),
                    E.Inventory(
                        E.SKU(v.amazon_sku),
                        E.FulfillmentCenterID(p.amazon_fulfillment_center_id),
                        E.Quantity(str(int(v.amazon_quantity))),
                        E.SwitchFulfillmentTo(p.amazon_fulfillment)
                    )
                )
            )
        try:
            response = Feed.submit_feed(
                    ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING),
                    '_POST_INVENTORY_AVAILABILITY_DATA_'
            )
        except MWSError as e:
            raise UserError(_('An error occured when submitting the inventory feed: %s') % e.response)

        return response.parsed.FeedSubmissionInfo.FeedSubmissionId

    def _submit_price_feed(self, Feed):
        IrConfigParameter = self.env['ir.config_parameter']
        currency_id = IrConfigParameter.get_param('amazon_currency')
        currency = self.env['res.currency'].browse(int(currency_id)).name

        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('Price'))
        for v in self:
            p = v.product_tmpl_id
            root.append(
                E.Message(
                    E.MessageID(str(v.id)),
                    E.Price(
                        E.SKU(v.amazon_sku),
                        E.StandardPrice(str(v.amazon_fixed_price or v.list_price), currency=str(currency)),
                    )
                )
            )
        response = Feed.submit_feed(
            ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING),
            '_POST_PRODUCT_PRICING_DATA_'
        )

        return response.parsed.FeedSubmissionInfo.FeedSubmissionId

    # Batch processing, group multiple products in a same feed
    def _list_new_products(self):
        root = self.env['amazon.instance'].generate_root_element()
        root.append(E.MessageType('Product'))

        for variant in self:
            root.append(variant._list_new_product_xml())

        print(ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING))
        return ET.tostring(root, pretty_print=True, xml_declaration=True, encoding=ENCODING)

    def _list_new_product_xml(self):
        product_xml = E.Product()
        if self.amazon_sku:
            product_xml.append(E.SKU(self.amazon_sku))
        product_xml.append(
            E.StandardProductID(
                E.Type(self.amazon_product_type),
                E.Value(self.amazon_product_type_value)
            )
        )
        #TODO do we handle tax code? different for each country,
        #didn't do it now because need to hardcode every tax code...
        product_xml.append(E.ProductTaxCode('A_GEN_NOTAX'))

        return E.Message(
            E.MessageID(str(self.id)),
            E.OperationType('Update'),
            product_xml
        )

    def _find_corresponding_offer(self):

        AmazonProducts = self.env['amazon.instance'].amazon_execute('Products')
        marketplace_id = self.env['amazon.instance'].get_amazon_marketplaceid()
        try:
            response = AmazonProducts.get_my_price_for_sku(marketplace_id, [self.amazon_sku])
        except MWSError as e:
            raise UserError(_('An error occured when linking the product: %s') % e.response)
        if 'Product' in response.parsed:
            return response.parsed.Product.Offers
        else:
            raise UserError(response.parsed.Error.Message)
