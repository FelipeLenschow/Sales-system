#!/usr/bin/env python3
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook, Workbook
import ctypes
import platform


# Constants for UI scaling
BASE_WIDTH = 1920
BASE_HEIGHT = 1080
Version = "0.1.3"

def is_numlock_on():
    if platform.system() != 'Windows':
        return True  # Assume Num Lock is on for non-Windows systems
    hllDll = ctypes.WinDLL ("User32.dll")
    VK_NUMLOCK = 0x90
    return hllDll.GetKeyState(VK_NUMLOCK) & 1

def set_numlock(state=True):

    if platform.system() != 'Windows':
        return  # Do nothing for non-Windows systems

    hllDll = ctypes.WinDLL ("User32.dll")
    VK_NUMLOCK = 0x90
    KEYEVENTF_EXTENDEDKEY = 0x0001
    KEYEVENTF_KEYUP = 0x0002

    current_state = is_numlock_on()
    if current_state != state:
        print("pressing NumLock")
        # Simulate key press
        hllDll.keybd_event(VK_NUMLOCK, 0x45, KEYEVENTF_EXTENDEDKEY | 0, 0)
        hllDll.keybd_event(VK_NUMLOCK, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)

class ProductDatabase:
    def __init__(self, filepath='produtos.xlsx'):
        self.filepath = filepath
        self.load_products()

    def load_products(self):
        try:
            # Tentar carregar o arquivo com MultiIndex no cabeçalho (2 linhas)
            self.df = pd.read_excel(self.filepath, header=[0, 1], dtype=str)

            # **Validação: Verificar se o DataFrame está vazio**
            if self.df.empty:
                raise ValueError("O arquivo está vazio.")

            # **Validação: Garantir que o cabeçalho tem o formato MultiIndex esperado**
            if not isinstance(self.df.columns, pd.MultiIndex):
                raise ValueError("O arquivo não possui um cabeçalho MultiIndex com duas linhas.")

            # Limpar espaços em branco nos cabeçalhos
            self.df.columns = pd.MultiIndex.from_tuples(
                [(str(x[0]).strip(), str(x[1]).strip()) for x in self.df.columns]
            )

            # Identificar as lojas no nível 0 do MultiIndex (exceto 'Todas')
            self.shops = [shop for shop in self.df.columns.levels[0] if shop != 'Todas']

            # Converter tipos das colunas
            self.df[('Todas', 'Codigo de Barras')] = self.df[('Todas', 'Codigo de Barras')].astype(str)
            self.df[('Metadata', 'Excel Row')] = self.df.index + 3
            for shop in self.shops:
                self.df[(shop, 'Preco')] = pd.to_numeric(self.df[(shop, 'Preco')], errors='coerce')
                self.df[(shop, 'Promo Preco')] = pd.to_numeric(self.df[(shop, 'Promo Preco')], errors='coerce')
                self.df[(shop, 'Promo Quantidade')] = pd.to_numeric(self.df[(shop, 'Promo Quantidade')],
                                                                    errors='coerce')

        except FileNotFoundError:
            messagebox.showerror("Erro", f"Arquivo {self.filepath} não encontrado.")
            self.df = pd.DataFrame()
            self.shops = []  # Inicializa como lista vazia para evitar novos erros
        except ValueError as ve:
            messagebox.showerror("Erro", f"Falha ao carregar produtos: {ve}")
            self.df = pd.DataFrame()
            self.shops = []
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
            self.df = pd.DataFrame()
            self.shops = []

    def add_product(self, product_info, shop):
        try:
            wb = load_workbook(self.filepath)
            ws = wb.active
        except FileNotFoundError:
            wb = Workbook()
            ws = wb.active
            ws.title = "Produtos"
            headers = [
                ("Todas", "Codigo de Barras"),
                ("Todas", "Sabor"),
                ("Todas", "Categoria"),
                (shop, "Preco"),
                (shop, "Promo Preco"),
                (shop, "Promo Quantidade")
            ]
            for col, (header1, header2) in enumerate(headers, start=1):
                ws.cell(row=1, column=col, value=header1)
                ws.cell(row=2, column=col, value=header2)

        # Map headers to column indices
        header_map = {
            f"{ws.cell(row=1, column=col).value} {ws.cell(row=2, column=col).value}": col
            for col in range(1, ws.max_column + 1)
        }
        excel_row = product_info.get('indexExcel', None)
        if excel_row is None:
            # Encontrar a próxima linha vazia para o código de barras
            data_start_row = 3
            barcode_col = header_map.get("Todas Codigo de Barras")
            excel_row = data_start_row
            while ws.cell(row=excel_row, column=barcode_col).value:
                excel_row += 1

        # Adicionar os dados na mesma linha
        ws.cell(row=excel_row, column=header_map["Todas Codigo de Barras"], value=product_info['barcode'])
        ws.cell(row=excel_row, column=header_map["Todas Sabor"], value=product_info['sabor'])
        ws.cell(row=excel_row, column=header_map["Todas Categoria"], value=product_info['categoria'])

        # Adicionar Preço
        preco = product_info['preco']
        ws.cell(row=excel_row, column=header_map[f"{shop} Preco"], value=float(preco))

        # Adicionar Promo Preço, se disponível
        promo_preco = product_info.get('promo_preco', None)
        ws.cell(row=excel_row, column=header_map[f"{shop} Promo Preco"],
                value=float(promo_preco) if promo_preco is not None else "")

        # Adicionar Promo Quantidade, se disponível
        promo_qt = product_info.get('promo_qt', None)
        ws.cell(row=excel_row, column=header_map[f"{shop} Promo Quantidade"],
                value=int(promo_qt) if promo_qt is not None else "")

        # Salvar e recarregar produtos
        wb.save(self.filepath)
        self.load_products()

    def filter_products(self, search_term, shop):
        if self.df.empty:
            return pd.DataFrame()
        mask = (
                self.df['Todas', 'Categoria'].str.contains(search_term, case=False, na=False) |
                self.df['Todas', 'Sabor'].str.contains(search_term, case=False, na=False) |
                self.df[shop, 'Preco'].astype(str).str.contains(search_term, case=False, na=False)
        )
        return self.df[mask]

    def get_unique_values(self, column, shop=None):
        if shop:
            return sorted(self.df[shop, column].dropna().unique().astype(str).tolist())
        return sorted(self.df['Todas', column].dropna().unique().astype(str).tolist())

    def get_products_by_barcode_and_shop(self, barcode, shop):
        """Busca todos os produtos pelo código de barras para a sorveteria atual."""
        try:
            # Filtrar pelo código de barras
            products = self.df[self.df['Todas', 'Codigo de Barras'].str.strip() == barcode.strip()]

            # Filtrar produtos com preço definido na loja atual
            products = products[pd.notna(products[(shop, 'Preco')])]

            if not products.empty:
                return products
            else:
                return pd.DataFrame()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar produtos: {e}")
            return pd.DataFrame()

    def get_products_by_barcode(self, barcode):
        """Retorna todos os produtos com o código de barras especificado em qualquer loja."""
        return self.df[self.df['Todas', 'Codigo de Barras'].str.strip() == barcode.strip()]

class Sale:
    def __init__(self, product_db, shop, payment_method="Débito"):
        self.product_db = product_db
        self.shop = shop
        self.payment_method = payment_method
        self.current_sale = {}
        self.final_price = 0.0

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

            if self.payment_method in ['Pix', 'Dinheiro'] and category_quantities[category] >= promo_qty:
                total_price += promo_price * quantity
            else:
                total_price += price * quantity

        self.final_price = total_price
        return self.final_price

    def add_product(self, product):
        excel_row = product[('Metadata', 'Excel Row')]
        if excel_row not in self.current_sale:
            self.current_sale[excel_row] = {
                'categoria': product[('Todas', 'Categoria')],
                'sabor': product[('Todas', 'Sabor')],
                'preco': product[(self.shop, 'Preco')],  # Já é float
                'promo_preco': product.get((self.shop, 'Promo Preco'), product[(self.shop, 'Preco')]),  # float
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


class POSApplication:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()  # Hide the root window initially

        # Ensure Num Lock is always on
        set_numlock(True)
        self.root.bind_all("<Num_Lock>", lambda event: (set_numlock(state=True), "break")[1])

        # Screen scaling
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.scale_factor = min(self.screen_width / BASE_WIDTH, self.screen_height / BASE_HEIGHT)

        # Initialize product database
        self.product_db = ProductDatabase()

        # Selected shop variable
        self.selected_shop_var = tk.StringVar()

        # Initialize payment method variable here
        self.payment_method_var = tk.StringVar(value="Débito")  # Initialized here

        # Initialize sale
        self.sale = None

        # Dictionary to keep track of widgets for each product
        self.product_widgets = {}

        # Initialize UI components
        self.initialize_ui()

    def initialize_ui(self):
        self.select_shop_window()

    def select_shop_window(self):
        def on_shop_select():
            selected = shop_combobox.get().strip()
            if selected:
                self.selected_shop_var.set(selected)
                shop_window.destroy()
                self.sale = Sale(self.product_db, selected, self.payment_method_var.get())
                self.build_main_window()
            else:
                messagebox.showerror("Erro", "Selecione a loja para continuar.")

        shop_window = tk.Toplevel(self.root)
        shop_window.title("Selecione a loja")
        shop_window.configure(bg="#1a1a2e")
        shop_window.attributes("-topmost", True)

        # Scale geometry
        win_width = int(400 * self.scale_factor)
        win_height = int(200 * self.scale_factor)
        shop_window.geometry(f"{win_width}x{win_height}")
        shop_window.resizable(False, False)
        shop_window.grab_set()  # Make modal

        # Center the window
        shop_window.update_idletasks()
        x = (shop_window.winfo_screenwidth() // 2) - (win_width // 2)
        y = (shop_window.winfo_screenheight() // 2) - (win_height // 2)
        shop_window.geometry(f"+{x}+{y}")

        # Fonts
        title_font = ("Arial", int(20 * self.scale_factor), "bold")
        version_font = ("Arial", int(8 * self.scale_factor))
        combobox_font = ("Arial", int(14 * self.scale_factor))

        # Label
        tk.Label(shop_window, text="Selecione a loja", bg="#1a1a2e", fg="#ffffff",
                 font=title_font).pack(pady=int(10 * self.scale_factor))

        tk.Label(shop_window, text=f"Versao: {Version}", bg="#1a1a2e", fg="#ffffff",
                 font=version_font).pack(pady=0)


        # Combobox
        shop_combobox = ttk.Combobox(
            shop_window,
            values=self.product_db.shops,  # Usando as lojas definidas no ProductDatabase
            state="readonly",
            font=combobox_font,
            width=20
        )
        shop_combobox.pack(pady=int(10 * self.scale_factor))
        shop_combobox.focus()

        # Select Button
        select_btn = ttk.Button(shop_window, text="Selecionar", command=on_shop_select)
        select_btn.pack(pady=int(20 * self.scale_factor))

        self.root.wait_window(shop_window)

    def minimize_application(self):
        """Minimiza a janela principal."""
        self.root.iconify()

    def build_main_window(self):
        self.root.deiconify()
        self.root.title("Sorveteria Lolla")
        self.root.attributes('-fullscreen', True)
        self.root.configure(bg="#1a1a2e")

        # Fonts
        title_font = ("Arial", int(50 * self.scale_factor), "bold")
        shop_font = ("Arial", int(32 * self.scale_factor))
        entry_font = ("Arial", int(16 * self.scale_factor))
        product_list_font = ("Arial", int(16 * self.scale_factor))
        label_font = ("Arial", int(16 * self.scale_factor))
        button_font = ("Arial", int(14 * self.scale_factor))
        final_price_font = ("Arial", int(50 * self.scale_factor), "bold")

        # Title Label
        title_label = ttk.Label(self.root, text="Sorveteria Lolla", font=title_font, background="#1a1a2e",
                                foreground="#ffffff")
        title_label.grid(row=0, column=0, columnspan=3, pady=int(20 * self.scale_factor), sticky="n")

        # Configure grid
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=0)
        self.root.grid_columnconfigure(2, weight=1)

        # Frame para Botões de Controle (Minimizar e Fechar)
        control_frame = tk.Frame(self.root, bg="#1a1a2e")
        control_frame.grid(row=0, column=2, sticky="ne", padx=int(10 * self.scale_factor), pady=int(10 * self.scale_factor))

        # Fechar Botão
        close_button = tk.Button(
            control_frame, text="✖", command=self.close_application, bg="#1a1a2e", fg="red",
            font=("Arial", int(16 * self.scale_factor)), borderwidth=0, width=3
        )
        close_button.pack(side=tk.RIGHT, padx=(0, 5))

        # Minimizar Botão
        minimize_button = tk.Button(
            control_frame, text="—", command=self.minimize_application, bg="#1a1a2e", fg="#ffffff",
            font=("Arial", int(16 * self.scale_factor)), borderwidth=0, width=3
        )
        minimize_button.pack(side=tk.RIGHT)

        # Selected Shop Label
        selected_shop_label = ttk.Label(
            self.root, text=f"{self.selected_shop_var.get()}", background="#1a1a2e",
            foreground="#ffffff", font=shop_font
        )
        selected_shop_label.grid(
            row=1, column=0, columnspan=3, padx=int(10 * self.scale_factor),
            pady=0, sticky="n"
        )

        # Barcode Entry
        self.barcode_entry = ttk.Entry(self.root, font=entry_font, width=30)
        self.barcode_entry.grid(
            row=2, column=0, columnspan=3, padx=int(10 * self.scale_factor),
            pady=int(5 * self.scale_factor)
        )
        self.barcode_entry.bind('<Return>', self.handle_barcode)

        # Add Product Button
        add_product_button = ttk.Button(
            self.root, text="Pesquisar produto", command=self.open_manual_add_window,
            style="Rounded.TButton", width=20
        )
        add_product_button.grid(
            row=3, column=0, columnspan=3, padx=int(10 * self.scale_factor),
            pady=int(5 * self.scale_factor)
        )

        # Sale Frame
        self.sale_frame = tk.Frame(self.root, bg="#1a1a2e")
        self.sale_frame.grid(
            row=4, column=0, padx=int(10 * self.scale_factor),
            pady=int(10 * self.scale_factor), sticky="nw"
        )

        # Final Price Label
        self.final_price_label = ttk.Label(
            self.root, text="R$0.00", font=final_price_font, background="#1a1a2e",
            foreground="#ffffff"
        )
        self.final_price_label.grid(
            row=2, column=2, columnspan=3, pady=int(20 * self.scale_factor),
            padx=int(50 * self.scale_factor), sticky="ne"
        )

        # Payment Method
        payment_method_label = ttk.Label(
            self.root, text="Método de pagamento:", background="#1a1a2e",
            foreground="#ffffff", font=label_font
        )
        payment_method_label.grid(
            row=3, column=2, padx=int(80 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="ne"
        )

        payment_method_combobox = ttk.Combobox(
            self.root, textvariable=self.payment_method_var,
            values=["Débito", "Pix", "Dinheiro", "Crédito"],
            state="readonly", font=entry_font, width=20
        )
        payment_method_combobox.grid(
            row=3, column=2, padx=int(60 * self.scale_factor),
            pady=int(35 * self.scale_factor), sticky="ne"
        )
        payment_method_combobox.current(0)
        payment_method_combobox.bind("<<ComboboxSelected>>", self.update_payment_method)

        # Style for Rounded Buttons
        style = ttk.Style()
        style.configure(
            "Rounded.TButton", font=button_font, padding=int(15 * self.scale_factor),
            relief="flat", background="#ffffff"
        )
        style.map(
            "Rounded.TButton",
            background=[("active", "#ffffff")],
            relief=[("pressed", "sunken")]
        )

        # Finalize and Clear Buttons
        button_frame = tk.Frame(self.root, bg="#1a1a2e")
        button_frame.grid(
            row=4, column=2, pady=int(10 * self.scale_factor),
            padx=int(50 * self.scale_factor), sticky="ne"
        )

        finalize_sale_button = ttk.Button(
            button_frame, text="Finalizar compra", style="Rounded.TButton",
            command=self.finalize_sale, width=20
        )
        finalize_sale_button.grid(
            row=1, column=0, padx=int(10 * self.scale_factor),
            pady=int(5 * self.scale_factor)
        )

        # Troco Frame
        troco_frame = tk.Frame(button_frame, bg="#1a1a2e")
        troco_frame.grid(row=2, column=0, pady=int(20 * self.scale_factor), sticky="ne")

        valor_pago_label = ttk.Label(
            troco_frame, text="Valor Pago:", background="#1a1a2e",
            foreground="#ffffff", font=label_font
        )
        valor_pago_label.grid(
            row=0, column=0, padx=int(5 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="e"
        )

        self.valor_pago_entry = ttk.Entry(
            troco_frame, font=entry_font, width=10
        )
        self.valor_pago_entry.grid(
            row=0, column=1, padx=int(5 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="w"
        )
        self.valor_pago_entry.bind("<KeyRelease>", self.calcular_troco)

        self.troco_label = ttk.Label(
            troco_frame, text="", background="#1a1a2e",
            foreground="#ffffff", font=label_font
        )
        self.troco_label.grid(
            row=1, column=0, columnspan=2, padx=int(5 * self.scale_factor),
            pady=int(5 * self.scale_factor)
        )

        self.sugestao_troco_label = ttk.Label(
            troco_frame, text="", background="#1a1a2e",
            foreground="#ffffff", font=label_font, justify="left"
        )
        self.sugestao_troco_label.grid(
            row=2, column=0, columnspan=2, padx=int(5 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="w"
        )

        # Clear Sale Button
        clear_sale_button = ttk.Button(
            self.root, text="Limpar", style="Rounded.TButton",
            command=self.new_sale, width=20
        )
        clear_sale_button.grid(
            row=5, column=2, padx=int(10 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="se"
        )

        # Add New Product Button
        add_new_product_button = ttk.Button(
            self.root, text="Adicionar novo produto", command=self.open_add_product_window,
            style="Rounded.TButton", width=20
        )
        add_new_product_button.grid(
            row=4, column=2, padx=int(10 * self.scale_factor),
            pady=int(5 * self.scale_factor), sticky="se"
        )

        self.update_sale_display()
        self.root.grid_rowconfigure(4, weight=1)

    def handle_barcode(self, event=None):
        barcode = self.barcode_entry.get().strip()
        if not barcode:
            return
        current_shop = self.selected_shop_var.get()

        # Obtém todos os produtos com o mesmo código de barras na loja atual
        matching_products = self.product_db.get_products_by_barcode_and_shop(barcode, current_shop)

        if matching_products.empty:
            # Verifica se existem produtos com o mesmo barcode em outras lojas
            other_products = self.product_db.get_products_by_barcode(barcode)
            if other_products.empty:
                # Nenhum produto encontrado, abrir janela para adicionar novo produto
                self.open_add_product_window(barcode=barcode)
            else:
                # Produtos encontrados em outras lojas, abrir janela para adicionar
                # Encontrar a primeira loja que possui o produto
                prefill_store = None
                for store in self.product_db.shops:
                    preco = other_products.iloc[0][(store, 'Preco')]
                    if pd.notna(preco):
                        prefill_store = store
                        break

                if prefill_store:
                    existing_product = other_products.iloc[0]
                    self.open_add_product_window(barcode=barcode, prefill_product=existing_product,
                                                 prefill_store=prefill_store)
                else:
                    # Se nenhuma loja possui o Preco definido, abrir sem pré-preenchimento
                    self.open_add_product_window(barcode=barcode)
        else:
            if len(matching_products) == 1:
                # Apenas um produto encontrado na loja atual, adiciona diretamente
                product = matching_products.iloc[0]
                self.sale.add_product(product)
                self.update_sale_display(product=product)  # Passa o produto completo
            else:
                # Múltiplos produtos encontrados na loja atual, abrir seleção
                self.select_product_window(matching_products)

        self.barcode_entry.delete(0, 'end')

    def select_product_window(self, matching_products):
        def on_select():
            selected_index = listbox.curselection()
            if selected_index:
                selected_product = matching_products.iloc[selected_index[0]]
                self.sale.add_product(selected_product)
                selection_window.destroy()
                self.update_sale_display(product=selected_product)  # Passa o produto completo

        selection_window = tk.Toplevel(self.root)
        selection_window.title("Selecione o Produto")
        selection_window.configure(bg="#8b0000")
        selection_window.attributes("-topmost", True)

        # Dimensionamento e posicionamento
        win_width = int(600 * self.scale_factor)
        win_height = int(400 * self.scale_factor)
        selection_window.geometry(f"{win_width}x{win_height}")
        selection_window.resizable(False, False)
        selection_window.grab_set()  # Torna a janela modal

        # Centralizar a janela
        selection_window.update_idletasks()
        x = (selection_window.winfo_screenwidth() // 2) - (win_width // 2)
        y = (selection_window.winfo_screenheight() // 2) - (win_height // 2)
        selection_window.geometry(f"+{x}+{y}")

        # Título
        title_font = ("Arial", int(20 * self.scale_factor), "bold")
        tk.Label(selection_window, text="Selecione o Produto", bg="#8b0000", fg="#ffffff", font=title_font).pack(
            pady=int(20 * self.scale_factor))

        # Listbox com scrollbar
        list_frame = tk.Frame(selection_window, bg="#8b0000")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, font=("Arial", int(14 * self.scale_factor)))
        listbox.pack(fill=tk.BOTH, expand=True)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)

        # Preencher a listbox com os produtos
        for idx, product in matching_products.iterrows():
            preco = product[(self.selected_shop_var.get(), 'Preco')]
            listbox.insert(tk.END,
                           f"{product[('Todas', 'Categoria')]}  {product[('Todas', 'Sabor')]}  R${preco:.2f}")

        # Botão de Seleção
        select_button = ttk.Button(selection_window, text="Selecionar", command=on_select)
        select_button.pack(pady=int(10 * self.scale_factor))

    def open_add_product_window(self, barcode=None, prefill_product=None, prefill_store=None):
        def parse_float(value):
            """Converte um valor para float, aceitando ',' ou '.' como separador decimal."""
            if value:
                try:
                    return float(value.replace(',', '.'))
                except ValueError:
                    raise ValueError(f"Valor inválido para número decimal: {value}")
            return None

        def parse_int(value):
            if value:
                try:
                    # Tenta converter para float primeiro
                    float_val = float(value.strip())
                    if float_val.is_integer():
                        return int(float_val)
                    else:
                        raise ValueError(f"Valor inválido para número inteiro: {value}")
                except ValueError:
                    raise ValueError(f"Valor inválido para número inteiro: {value}")
            return None

        def save_product():
            try:
                # Obter e limpar os valores dos campos
                barcode_val = barcode_entry.get().strip()
                sabor = flavor_combobox.get().strip()
                categoria = category_combobox.get().strip()
                preco = price_entry.get().strip()
                promo_preco = promo_price_entry.get().strip()
                promo_qt = promo_qty_entry.get().strip()

                shop = self.selected_shop_var.get().strip()

                # Validação dos campos obrigatórios
                if not all([barcode_val, sabor, categoria, preco, shop]):
                    messagebox.showerror("Erro", "Preencha todos os campos necessários")
                    return

                # Conversão para valores numéricos, permitindo ',' ou '.'
                preco_val = parse_float(preco)
                promo_preco_val = parse_float(promo_preco) if promo_preco else None
                promo_qt_val = parse_int(promo_qt) if promo_qt else None

                # Construir o dicionário 'product_info'
                product_info = {
                    'indexExcel': index_value,
                    'barcode': barcode_val,
                    'sabor': sabor,
                    'categoria': categoria,
                    'preco': preco_val,
                    'promo_preco': promo_preco_val,
                    'promo_qt': promo_qt_val,
                }

                # Adicionar ou atualizar o produto na base de dados
                self.product_db.add_product(product_info, shop)
                self.update_sale_display()  # Atualizar a exibição da venda
                add_product_window.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao adicionar/atualizar o produto: {e}")

        # Criar a janela de adicionar produto
        add_product_window = tk.Toplevel(self.root)
        add_product_window.title("Adicionar/Atualizar Produto")
        add_product_window.configure(bg="#8b0000")
        add_product_window.attributes("-topmost", True)

        # Frame para inputs
        input_frame = tk.Frame(add_product_window, bg="#8b0000")
        input_frame.pack(pady=10, padx=10)

        # Código de Barras
        tk.Label(input_frame, text="Código de Barras", bg="#8b0000", fg="#ffffff").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        barcode_entry = ttk.Entry(input_frame)
        barcode_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        if barcode:
            barcode_entry.insert(0, barcode)

        # Preencher campos com valores existentes, se houver
        if prefill_product is not None:
            index_value = prefill_product[('Metadata', 'Excel Row')]
            sabor_value = prefill_product[('Todas', 'Sabor')]
            categoria_value = prefill_product[('Todas', 'Categoria')]
            if prefill_store:
                preco_value = prefill_product.get((prefill_store, 'Preco'), "")
                promo_preco_value = prefill_product.get((prefill_store, 'Promo Preco'), "")
                promo_qt_value = prefill_product.get((prefill_store, 'Promo Quantidade'), "")
            else:
                preco_value = ""
                promo_preco_value = ""
                promo_qt_value = ""
        else:
            index_value = None
            sabor_value = ""
            categoria_value = ""
            preco_value = ""
            promo_preco_value = ""
            promo_qt_value = ""

        # Sabor
        tk.Label(input_frame, text="Sabor:", bg="#8b0000", fg="#ffffff").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        flavor_combobox = ttk.Combobox(
            input_frame,
            values=self.product_db.get_unique_values('Sabor'),
            state="normal"
        )
        flavor_combobox.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        flavor_combobox.insert(0, sabor_value)  # Preencher sabor se disponível

        # Categoria
        tk.Label(input_frame, text="Categoria:", bg="#8b0000", fg="#ffffff").grid(row=0, column=4, padx=5, pady=5, sticky="e")
        category_combobox = ttk.Combobox(
            input_frame,
            values=self.product_db.get_unique_values('Categoria'),
            state="normal"
        )
        category_combobox.grid(row=0, column=5, padx=5, pady=5, sticky="w")
        category_combobox.insert(0, categoria_value)  # Preencher categoria se disponível

        # Preço
        tk.Label(input_frame, text="Preço:", bg="#8b0000", fg="#ffffff").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        price_entry = ttk.Entry(input_frame)
        price_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        if preco_value and preco_value != "":
            price_entry.insert(0, preco_value)

        # Promo Preço
        tk.Label(input_frame, text="Promo Preço:", bg="#8b0000", fg="#ffffff").grid(row=1, column=2, padx=5, pady=5, sticky="e")
        promo_price_entry = ttk.Entry(input_frame)
        promo_price_entry.grid(row=1, column=3, padx=5, pady=5, sticky="w")
        if promo_preco_value and promo_preco_value != "":
            promo_price_entry.insert(0, promo_preco_value)

        # Promo Quantidade
        tk.Label(input_frame, text="Promo Quantidade:", bg="#8b0000", fg="#ffffff").grid(row=1, column=4, padx=5, pady=5, sticky="e")
        promo_qty_entry = ttk.Entry(input_frame)
        promo_qty_entry.grid(row=1, column=5, padx=5, pady=5, sticky="w")
        if promo_qt_value and promo_qt_value != "":
            promo_qty_entry.insert(0, promo_qt_value)

        # Botão para salvar
        save_button = ttk.Button(add_product_window, text="Salvar/Atualizar Produto", command=save_product)
        save_button.pack(pady=20)

    def open_manual_add_window(self):
        def search_products():
            product_listbox.delete(0, tk.END)
            if search_entry.get() != "":
                search_term = search_entry.get().lower()
                self.filtered_products = self.product_db.filter_products(search_term, self.selected_shop_var.get())

                for _, product in self.filtered_products.iterrows():
                    preco = product[(self.selected_shop_var.get(), 'Preco')]
                    product_listbox.insert(
                        tk.END,
                        f"{product['Todas', 'Sabor']} ({product['Todas', 'Categoria']}) - R${preco:.2f}"
                    )


        def on_select():
            selected_index = product_listbox.curselection()
            if selected_index:
                selected_product = self.filtered_products.iloc[selected_index[0]]
                self.sale.add_product(selected_product)
                manual_add_window.destroy()
                self.update_sale_display()

        manual_add_window = tk.Toplevel(self.root)
        manual_add_window.title("Adicionar produto manualmente")
        manual_add_window.configure(bg="#8b0000")
        manual_add_window.attributes("-topmost", True)

        search_frame = tk.Frame(manual_add_window, bg="#8b0000")
        search_frame.pack(padx=10, pady=10)

        search_label = ttk.Label(
            search_frame, text="Pesquisar por categoria, sabor, ou preco:",
            background="#8b0000", foreground="#ffffff"
        )
        search_label.pack(side=tk.LEFT)

        search_entry = ttk.Entry(search_frame)
        search_entry.pack(side=tk.LEFT, padx=(5, 0))
        search_entry.bind("<Return>", lambda event: (search_products(), "break")[1])
        search_entry.bind("<KeyRelease>", lambda event: (search_products(), "break")[1])


        search_button = ttk.Button(search_frame, text="Pesquisar", command=search_products)
        search_button.pack(side=tk.LEFT, padx=(5, 0))

        product_listbox = tk.Listbox(manual_add_window, height=10, bg="#ffffff", font=("Arial", 12))
        product_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        select_button = ttk.Button(manual_add_window, text="Selecionar", command=on_select)
        select_button.pack(pady=5)

    def update_payment_method(self, event=None):
        self.sale.payment_method = self.payment_method_var.get()
        self.sale.apply_promotion()
        self.update_sale_display()

    def create_or_update_product_widget(self, excel_row, details):

        if excel_row not in self.product_widgets:
            row = len(self.product_widgets)

            # Nome e categoria
            text_widget = tk.Text(
                self.sale_frame, height=1, width=35, bg="#1a1a2e", fg="#ffffff",
                font=("Arial", 18), bd=0, highlightthickness=0
            )
            text_widget.grid(row=row, column=0, padx=50, pady=2, sticky="w")
            text_widget.insert(tk.END, f"{details['categoria']} - {details['sabor']}")
            text_widget.tag_configure("bold", font=("Arial", 18, "bold"))
            text_widget.config(state=tk.DISABLED)

            # Label de preço
            price_label = tk.Label(
                self.sale_frame, text="", bg="#1a1a2e",
                fg="#ffffff", font=("Arial", 18)
            )
            price_label.grid(row=row, column=2, padx=5, pady=2)

            # Entry de quantidade
            quantity_var = tk.StringVar(value=str(details['quantidade']))
            quantity_entry = ttk.Entry(
                self.sale_frame, textvariable=quantity_var, width=5, font=("Arial", 18)
            )
            quantity_entry.grid(row=row, column=1, padx=5, pady=2)
            quantity_entry.bind("<KeyRelease>", lambda event: self.update_quantity_dynamic(excel_row, quantity_var))
            quantity_entry.bind("<FocusIn>", self.select_all_text)

            # Botão de remover
            delete_button = tk.Button(
                self.sale_frame, text="✖",
                command=lambda b=excel_row: self.delete_product(b),
                bg="#1a1a2e", fg="#ffffff", font=("Arial", int(16 * self.scale_factor)),
                borderwidth=0
            )
            delete_button.grid(row=row, column=3, padx=5, pady=2)

            # Botão de editar
            edit_button = tk.Button(
                self.sale_frame, text="✎",
                command=lambda b=excel_row: self.edit_product(b),
                bg="#1a1a2e", fg="#ffffff", font=("Arial", int(16 * self.scale_factor)),
                borderwidth=0
            )
            edit_button.grid(row=row, column=4, padx=5, pady=2)

            # Armazena os widgets no dicionário
            self.product_widgets[excel_row] = {
                'text_widget': text_widget,
                'price_label': price_label,
                'quantity_entry': quantity_entry,
                'quantity_var': quantity_var,
                'delete_button': delete_button,
                'edit_button': edit_button
            }
        else:
            # Atualiza widgets existentes
            widgets = self.product_widgets[excel_row]
            widgets['text_widget'].config(state=tk.NORMAL)
            widgets['text_widget'].delete("1.0", tk.END)
            widgets['text_widget'].insert(tk.END, f"{details['categoria']} - {details['sabor']}")
            widgets['text_widget'].config(state=tk.DISABLED)

            # Atualiza a quantidade
            widgets['quantity_var'].set(str(details['quantidade']))

        # Atualiza o preço com base na promoção
        widgets = self.product_widgets[excel_row]
        if (self.sale.payment_method in ['Pix', 'Dinheiro'] and
                self.category_quantities.get(details['categoria'], 0) >= details['promo_qt']):
            price = details['promo_preco']
            fg_color = "#88ff88"
        else:
            price = details['preco']
            fg_color = "#ffffff"
        widgets['price_label'].config(text=f"R${price:.2f}", fg=fg_color)

    def update_sale_display(self, product=None):
        # Aplica promoções e calcula o preço final
        final_price = self.sale.apply_promotion()
        self.final_price_label.config(text=f"R${final_price:.2f}")

        # Calcula quantidades por categoria
        self.category_quantities = {}
        for details in self.sale.current_sale.values():
            category = details['categoria']
            quantity = details['quantidade']
            self.category_quantities[category] = self.category_quantities.get(category, 0) + quantity

        # Atualiza widgets existentes ou cria novos
        for excel_row in self.sale.current_sale.keys():
            try:
                # Buscar os detalhes atualizados do banco de dados
                product_series = self.product_db.df.loc[excel_row - 3]  # Ajustar para índice do DataFrame
                details = {
                    'categoria': product_series[('Todas', 'Categoria')],
                    'sabor': product_series[('Todas', 'Sabor')],
                    'preco': float(product_series[(self.sale.shop, 'Preco')]),
                    'promo_preco': float(product_series[(self.sale.shop, 'Promo Preco')]) if pd.notna(
                        product_series[(self.sale.shop, 'Promo Preco')]) else float(
                        product_series[(self.sale.shop, 'Preco')]),
                    'promo_qt': int(product_series[(self.sale.shop, 'Promo Quantidade')]) if pd.notna(
                        product_series[(self.sale.shop, 'Promo Quantidade')]) else 1,
                    'quantidade': self.sale.current_sale[excel_row]['quantidade'],
                    'indexExcel': excel_row
                }
                self.create_or_update_product_widget(excel_row, details)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao atualizar produto {excel_row}: {e}")

        # Remove widgets que não estão mais na venda
        existing_excel_rows = set(self.product_widgets.keys())
        current_excel_rows = set(self.sale.current_sale.keys())
        for excel_row in existing_excel_rows - current_excel_rows:
            widgets = self.product_widgets[excel_row]
            for widget in widgets.values():
                if isinstance(widget, (tk.Widget, ttk.Widget)):
                    widget.destroy()
            del self.product_widgets[excel_row]

        # Restaurar o foco no código de barras
        self.root.bind("<Return>", lambda event: (self.barcode_entry.focus(), "break")[1])
        self.root.bind("<F12>", lambda event: (self.barcode_entry.focus(), "break")[1])
        self.root.bind("<End>", lambda event: (self.barcode_entry.focus(), "break")[1])

        # Se um produto foi passado, focar no widget de quantidade correspondente
        if product is not None:
            excel_row = product[('Metadata', 'Excel Row')]
            if excel_row in self.product_widgets:
                self.product_widgets[excel_row]['quantity_entry'].focus_set()

        if self.valor_pago_entry.get():
            self.calcular_troco()

    def select_all_text(self, event):
        event.widget.select_range(0, 'end')
        event.widget.icursor('end')
        return 'break'

    def update_quantity_dynamic(self, index_excel, quantity_var):
        try:
            new_quantity = int(quantity_var.get())
            if new_quantity < 0:
                new_quantity = 0
            self.sale.update_quantity(index_excel, new_quantity)
            # Update only the price label for this product
            self.update_sale_display()
        except ValueError:
            if quantity_var.get() != "" :
                messagebox.showerror("Quantidade Inválida", f"Por favor, insira um número válido. ,{quantity_var.get()},")

    def delete_product(self, excel_row):
        self.sale.remove_product(excel_row)
        if excel_row in self.product_widgets:
            # Destrói todos os widgets associados ao produto
            for widget in self.product_widgets[excel_row].values():
                if isinstance(widget, (tk.Widget, ttk.Widget)):
                    widget.destroy()
            # Remove o produto do dicionário
            del self.product_widgets[excel_row]
        self.update_sale_display()

    def edit_product(self, index_excel):
        """Abre uma janela para editar as informações do produto no arquivo Excel."""
        # Encontrar o índice do Excel para este produto
        excel_row = self.sale.current_sale[index_excel]['indexExcel']
        product_series = self.product_db.df.loc[excel_row - 3]  # Ajustar para índice do DataFrame

        # Obter os dados atuais do produto
        current_barcode = product_series[('Todas', 'Codigo de Barras')]
        current_sabor = product_series[('Todas', 'Sabor')]
        current_categoria = product_series[('Todas', 'Categoria')]
        current_preco = product_series[(self.sale.shop, 'Preco')]
        current_promo_preco = product_series[(self.sale.shop, 'Promo Preco')]
        current_promo_qt = product_series[(self.sale.shop, 'Promo Quantidade')]

        def save_changes():
            try:
                # Obter e limpar os valores dos campos
                new_barcode = barcode_entry.get().strip()
                new_sabor = sabor_entry.get().strip()
                new_categoria = categoria_entry.get().strip()
                new_preco = preco_entry.get().strip()
                new_promo_preco = promo_preco_entry.get().strip()
                new_promo_qt = promo_qt_entry.get().strip()

                # Validação dos campos obrigatórios
                if not all([new_barcode, new_sabor, new_categoria, new_preco]):
                    messagebox.showerror("Erro", "Preencha todos os campos necessários")
                    return

                # Conversão para valores numéricos, permitindo ',' ou '.'
                new_preco_val = parse_float(new_preco)
                new_promo_preco_val = parse_float(new_promo_preco) if new_promo_preco else None
                new_promo_qt_val = parse_int(new_promo_qt) if new_promo_qt else None

                # Construir o dicionário 'product_info' para atualização
                product_info = {
                    'indexExcel': excel_row,
                    'barcode': new_barcode,
                    'sabor': new_sabor,
                    'categoria': new_categoria,
                    'preco': new_preco_val,
                    'promo_preco': new_promo_preco_val,
                    'promo_qt': new_promo_qt_val,
                }

                # Adicionar ou atualizar o produto na base de dados
                self.product_db.add_product(product_info, self.sale.shop)

                # Atualizar os detalhes na venda atual, se o produto estiver na venda
                if excel_row in self.sale.current_sale:
                    self.sale.current_sale[excel_row].update({
                        'categoria': new_categoria,
                        'sabor': new_sabor,
                        'preco': new_preco_val,
                        'promo_preco': new_promo_preco_val if new_promo_preco_val is not None else new_preco_val,
                        'promo_qt': new_promo_qt_val if new_promo_qt_val is not None else 1,
                    })


                self.update_sale_display()  # Atualizar a exibição da venda
                edit_window.destroy()
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao atualizar o produto: {e}")

        def parse_float(value):
            """Converte um valor para float, aceitando ',' ou '.' como separador decimal."""
            if value:
                try:
                    return float(value.replace(',', '.'))
                except ValueError:
                    raise ValueError(f"Valor inválido para número decimal: {value}")
            return None

        def parse_int(value):
            """Converte um valor para int, aceitando valores que são floats com .0"""
            if value:
                try:
                    # Tenta converter para float primeiro
                    float_val = float(value.strip())
                    if float_val.is_integer():
                        return int(float_val)
                    else:
                        raise ValueError(f"Valor inválido para número inteiro: {value}")
                except ValueError:
                    raise ValueError(f"Valor inválido para número inteiro: {value}")
            return None

        # Criar a janela de edição
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Editar Produto")
        edit_window.configure(bg="#1a1a2e")

        # Frame para inputs
        input_frame = tk.Frame(edit_window, bg="#1a1a2e")
        input_frame.pack(pady=10, padx=10)

        # Código de Barras
        tk.Label(input_frame, text="Código de Barras:", bg="#1a1a2e", fg="#ffffff").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        barcode_entry = ttk.Entry(input_frame)
        barcode_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        barcode_entry.insert(0, current_barcode)

        # Sabor
        tk.Label(input_frame, text="Sabor:", bg="#1a1a2e", fg="#ffffff").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        sabor_entry = ttk.Entry(input_frame)
        sabor_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        sabor_entry.insert(0, current_sabor)

        # Categoria
        tk.Label(input_frame, text="Categoria:", bg="#1a1a2e", fg="#ffffff").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        categoria_entry = ttk.Entry(input_frame)
        categoria_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        categoria_entry.insert(0, current_categoria)

        # Preço
        tk.Label(input_frame, text="Preço:", bg="#1a1a2e", fg="#ffffff").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        preco_entry = ttk.Entry(input_frame)
        preco_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        preco_entry.insert(0, current_preco)

        # Promo Preço
        tk.Label(input_frame, text="Promo Preço:", bg="#1a1a2e", fg="#ffffff").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        promo_preco_entry = ttk.Entry(input_frame)
        promo_preco_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        if pd.notna(current_promo_preco):
            promo_preco_entry.insert(0, current_promo_preco)

        # Promo Quantidade
        tk.Label(input_frame, text="Promo Quantidade:", bg="#1a1a2e", fg="#ffffff").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        promo_qt_entry = ttk.Entry(input_frame)
        promo_qt_entry.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        if pd.notna(current_promo_qt):
            promo_qt_entry.insert(0, current_promo_qt)

        # Botão para salvar alterações
        save_button = ttk.Button(edit_window, text="Salvar Alterações", command=save_changes)
        save_button.pack(pady=20)

    def calcular_troco(self, event=None):
        try:
            valor_pago = float(self.valor_pago_entry.get().replace(",", "."))
            troco = valor_pago - self.sale.final_price
            if troco < 0.0:
                self.troco_label.config(text=f"Troco: R${troco:.2f}", foreground="#ff5555")
            else:
                self.troco_label.config(text=f"Troco: R${troco:.2f}", foreground="#55ff55")
            self.sugestao_troco_label.config(text="")
        except ValueError:
            self.troco_label.config(text="")
            self.sugestao_troco_label.config(text="")

    def finalize_sale(self):
        if not self.sale.current_sale:
            messagebox.showerror("Erro", "Sem produtos nas vendas!")
            return

        # Apply promotion and calculate final price
        final_price = self.sale.apply_promotion()

        # Save sale details
        now = datetime.now()
        sale_data = {
            'Data': [now.strftime('%Y-%m-%d')],
            'Horario': [now.strftime('%H:%M:%S')],
            'Preco Final': [final_price],
            'Metodo de pagamento': [self.sale.payment_method],
            'Produtos': [self.sale.current_sale],
            'Quantidade de produtos': [sum(product['quantidade'] for product in self.sale.current_sale.values())]
        }
        sale_df = pd.DataFrame(sale_data)

        # Load or create sales history
        try:
            sales_history = pd.read_excel('Historico_vendas.xlsx')
        except FileNotFoundError:
            sales_history = pd.DataFrame(
                columns=['Data', 'Horario', 'Preco Final', 'Metodo de pagamento', 'Produtos', 'Quantidade de produtos']
            )

        # Append new sale
        updated_sales_history = pd.concat([sales_history, sale_df], ignore_index=True)
        updated_sales_history.to_excel('Historico_vendas.xlsx', index=False)

        # Reset sale
        self.new_sale()

    def new_sale(self):
        self.sale = Sale(self.product_db, self.selected_shop_var.get(), self.payment_method_var.get())
        self.update_sale_display()
        self.payment_method_var.set("Débito")
        self.valor_pago_entry.delete(0, 'end')
        self.troco_label.config(text="")
        self.sugestao_troco_label.config(text="")

        # Clear product widgets
        for widgets in self.product_widgets.values():
            for widget in widgets.values():
                if isinstance(widget, (tk.Widget, ttk.Widget)):  # Destrói apenas widgets Tkinter
                    widget.destroy()
        self.product_widgets.clear()

    def close_application(self):
        self.root.quit()


if __name__ == "__main__":
    root = tk.Tk()
    app = POSApplication(root)
    root.mainloop()
