import uuid

class Sale:
    def __init__(self, product_db, shop, payment_method=""):
        self.product_db = product_db
        self.shop = shop
        self.payment_method = payment_method
        self.current_sale = {}
        self.final_price = 0.0
        self.id = str(uuid.uuid4())

    def apply_promotion(self):
        total_price = 0.0
        category_quantities = {}

        # Soma as quantidades por categoria
        for product in self.current_sale.values():
            category = product['categoria']
            quantity = product['quantidade']
            category_quantities[category] = category_quantities.get(category, 0) + quantity

        # Calcula o preço total com promoções
        for product in self.current_sale.values():
            category = product['categoria']
            quantity = product['quantidade']
            promo_qty = product['promo_qt']
            price = product['preco']
            promo_price = product['promo_preco']

            if (self.payment_method in ['Pix', 'Dinheiro'] and
                    promo_qty is not None and
                    category_quantities[category] >= promo_qty):
                total_price += promo_price * quantity
            else:
                total_price += price * quantity

        self.final_price = total_price
        return self.final_price

    def add_product(self, product):
        excel_row = product[('Metadata', 'Excel Row')]
        if excel_row not in self.current_sale or (type(excel_row) == str and excel_row.startswith('Manual')):
            self.current_sale[excel_row] = {
                'categoria': product[('Todas', 'Categoria')],
                'sabor': product[('Todas', 'Sabor')],
                'preco': product[(self.shop, 'Preco')],
                'promo_preco': product.get((self.shop, 'Promo Preco'), product[(self.shop, 'Preco')]),
                'promo_qt': product.get((self.shop, 'Promo Quantidade'), 1),  # int
                'quantidade': 1,
                'indexExcel': excel_row
            }
        else:
            self.current_sale[excel_row]['quantidade'] += 1

    def remove_product(self, excel_row):
        if excel_row in self.current_sale:
            del self.current_sale[excel_row]

    def update_quantity(self, excel_row, quantity):
        if excel_row in self.current_sale:
            self.current_sale[excel_row]['quantidade'] = max(quantity, 0)
            if self.current_sale[excel_row]['quantidade'] == 0:
                self.remove_product(excel_row)
