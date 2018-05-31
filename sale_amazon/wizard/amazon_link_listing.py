# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from datetime import datetime

from odoo import models, fields, api, _
from odoo.exceptions import UserError


class amazon_link_listing(models.TransientModel):
    _name = 'amazon.link.listing'

    @api.model
    def default_get(self, fields):
        res = super(amazon_link_listing, self).default_get(fields)

        active_id = self._context.get('active_id', [])
        active_model = self._context.get('active_model', False)

        lines = []
        for rec in self.env[active_model].browse(active_id).product_variant_ids:
            lines.append((0, 0, {
                'name': rec.display_name,
                'amazon_sku': rec.amazon_sku,
                'product_id': rec.id,
            }))
        res['product_variant_ids'] = lines
        res['product_tmpl_id'] = active_id
        return res

    amazon_id_type = fields.Selection([('ASIN', 'ASIN'), ('ISBN', 'ISBN'), ('UPC', 'UPC'), ('EAN', 'EAN')], string='Amazon ID Type', required=True, help="The type of product identifier that Id values refer to.")
    amazon_id = fields.Char('Amazon ID', required=True)
    amazon_sku = fields.Char('Amazon SKU', required=True)
    product_tmpl_id = fields.Many2one('product.template')
    product_variant_count = fields.Integer(compute="_get_product_variant_count")
    product_variant_ids = fields.One2many("amazon.link.listing.variant", "link_listing_id")

    @api.multi
    @api.depends('product_tmpl_id', 'product_variant_ids')
    def _get_product_variant_count(self):
        for rec in self:
            rec.product_variant_count = len(rec.product_tmpl_id.attribute_line_ids)

    @api.one
    def link_listing(self):

        # TODO: add quantity, amazon tax?, price, condition shipping cost from offer

        # TODO: retrieve product category (get_product_categories_for_sku)

        product = self.env['product.template'].browse(self._context.get('active_id'))

        if self.product_variant_count and self.product_variant_ids.filtered(lambda x: not x.amazon_sku):
            raise UserError(_("Please, Enter all variant Amazon SKU."))

        product.amazon_product_type = self.amazon_id_type
        product.amazon_product_type_value = self.amazon_id
        product.amazon_sku = self.amazon_sku

        found_product = product._find_corresponding_products()

        if not found_product:
            raise UserError(_("No associated product found."))
        elif not self.amazon_sku and len(found_product) > 1:
            raise UserError(_("More than one associated product found."))

        found_product = found_product.Product

        product_values = {}

        product_values['amazon_last_sync'] = datetime.now()
        if not self.product_variant_count:
            product_values['amazon_product_type'] = "ASIN"
            product_values['amazon_product_type_value'] = found_product.Identifiers.MarketplaceASIN.ASIN

        # VARIANTS
        relationships = found_product.Relationships
        if relationships:
            print(relationships)
            if 'VariationParent' in relationships:
                raise UserError(_("This product seems to be a product variation. Link the main product instead."))
            elif 'VariationChild' in relationships:
                variant_child = relationships.VariationChild
                if not isinstance(variant_child, list):
                    variant_child = [relationships.VariationChild]
                if len(variant_child) != product.product_variant_count:
                    raise UserError(_('Number of variants differs between Amazon product and Odoo product.'))
                for amazon_variant in variant_child:
                    # Variations seem ok, match with odoo product
                    amazon_variant_asin = amazon_variant.pop('Identifiers').MarketplaceASIN.ASIN
                    specs = amazon_variant.keys()
                    attrs = []
                    for spec in specs:
                        attr = self.env['product.attribute.value'].search([('name', '=', amazon_variant[spec]['value'])])
                        attrs.append(('attribute_value_ids', '=', attr.id))
                    variant = self.env['product.product'].search(attrs).filtered(
                        lambda l: l.product_tmpl_id.id == product.id)
                    variant.write({
                        'amazon_use': True,
                        'amazon_product_type': 'ASIN',
                        'amazon_product_type_value': amazon_variant_asin,
                        'amazon_sku': self.product_variant_ids.filtered(lambda x: x.product_id.id == variant.id).amazon_sku
                    })
            else:
                raise UserError(_('Type of relationship unknown (only variations are possible)'))
        else:
            if product.product_variant_count > 1:
                raise UserError(_('No relationship found on the Amazon product but the Odoo product have variants.'))
            elif product.product_variant_count == 0:
                raise UserError(_('The Odoo product is a service and can\'t be sold on Amazon.'))
            else:
                variant = product.product_variant_ids
                variant.write({
                    'amazon_use': True,
                })

        product.write(product_values)

        # More information on product.product
        for variant in product.product_variant_ids:
            variant_values = {}
            found_offer = variant._find_corresponding_offer()
            if found_offer:
                found_offer = found_offer.Offer
                # Amazon sends two offers when the seller has both AFN and MFN fulfilment for the product
                if isinstance(found_offer, list):
                    for offer in found_offer:
                        if offer.SellerSKU == variant.amazon_sku:
                            found_offer = offer
                            break
                variant_values['amazon_fixed_price'] = found_offer.BuyingPrice.ListingPrice.Amount
                variant_values['amazon_shipping_price'] = found_offer.BuyingPrice.Shipping.Amount
                variant_values['amazon_condition'] = found_offer.ItemCondition
                variant_values['amazon_fulfillment'] = 'MFN' if found_offer.FulfillmentChannel == 'MERCHANT' else 'AFN'
                if variant_values['amazon_fulfillment'] == 'MFN':
                    product.product_variant_ids.write({'amazon_state': 'quantity_missing'})

                elif variant_values['amazon_fulfillment'] == 'AFN':
                    found_inventory = variant._list_inventory_supply_by_amazon()

                    if found_inventory:
                        found_inventory = found_inventory.member
                        # Maybe we should use 'InStockSupplyQuantity' ?
                        variant_values['amazon_quantity'] = found_inventory.TotalSupplyQuantity
                        variant_values['amazon_state'] = 'published'
                    else:
                        product.product_variant_ids.write({'amazon_state': 'quantity_missing'})
            else:
                variant_values['amazon_state'] = 'listing_ended'
            variant.write(variant_values)

class amazon_link_listing_variant(models.TransientModel):
    _name = 'amazon.link.listing.variant'

    name = fields.Char("Name")
    amazon_sku = fields.Char('Amazon SKU')
    product_id = fields.Many2one("product.product")
    link_listing_id = fields.Many2one("amazon.link.listing")
