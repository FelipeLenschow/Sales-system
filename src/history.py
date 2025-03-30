import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import math


class ToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tooltip = None
        self.current_item = None  # Track the currently hovered item
        self.widget.bind("<Motion>", self.on_motion)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def on_motion(self, event):
        # Check if the mouse is over a valid item
        item = self.widget.identify_row(event.y)
        if item != self.current_item:  # Only update if the hovered item changes
            self.current_item = item
            if item:
                # Get the 'Produtos' value for the hovered item
                produtos = self.widget.item(item, 'values')[4]
                if produtos:
                    try:
                        # Safely evaluate the string into a dictionary
                        produtos_dict = self.safe_eval_produtos(produtos)

                        # Format the products into a readable list
                        formatted_products = self.format_products(produtos_dict)

                        # Get the position of the mouse
                        bbox = self.widget.bbox(item)
                        if bbox:  # Check if the bounding box is valid
                            x, y, _, _ = bbox
                            x += self.widget.winfo_rootx() + 25
                            y += self.widget.winfo_rooty() + 25

                            # Update or create the tooltip
                            if self.tooltip:
                                # Update the existing tooltip
                                self.tooltip.wm_geometry(f"+{x}+{y}")
                                self.tooltip_label.config(text=formatted_products)
                            else:
                                # Create the tooltip window
                                self.tooltip = tk.Toplevel(self.widget)
                                self.tooltip.wm_overrideredirect(True)
                                self.tooltip.wm_geometry(f"+{x}+{y}")

                                # Add a label to display the formatted products
                                self.tooltip_label = tk.Label(self.tooltip, text=formatted_products, background="#ffffe0", relief="solid", borderwidth=1, justify=tk.LEFT)
                                self.tooltip_label.pack()
                    except Exception as e:
                        print(f"Error formatting products: {e}")
            else:
                # Hide the tooltip if the mouse is not over a valid item
                self.hide_tooltip()

    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
            self.current_item = None  # Reset the current item

    def safe_eval_produtos(self, produtos):
        # Define a safe namespace for evaluation
        safe_namespace = {
            "np": np,  # Allow numpy functions
            "int": int,  # Allow Python int
            "float": float,  # Allow Python float
            "nan": np.nan,  # Allow numpy nan
        }

        # Evaluate the string in the safe namespace
        return eval(produtos, {"__builtins__": {}}, safe_namespace)

    def format_products(self, produtos_dict):
        # Format the dictionary into a readable list
        formatted_text = "Produtos:\n"
        for product_id, details in produtos_dict.items():
            # Handle None or nan values in the dictionary
            for key, value in details.items():
                if value is None or (isinstance(value, float) and math.isnan(value)):
                    details[key] = "N/A"

            # Format the price with R$ and two decimal places
            preco = details['preco']
            if isinstance(preco, (int, float)):
                preco = f"R${preco:.2f}"
            else:
                preco = "R$0.00"

            formatted_text += (
                f"- {details['categoria']} ({details['sabor']}): "
                f"{details['quantidade']} unidade(s) a {preco}\n"
            )
        return formatted_text
class SalesHistoryWindow:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Histórico de Vendas")
        self.window.geometry("800x600")

        # Create a frame to hold the Treeview and scrollbars
        self.tree_frame = tk.Frame(self.window)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)

        # Create a Treeview widget
        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=('Data', 'Horario', 'Preco Final', 'Metodo de pagamento', 'Produtos'),
            show='headings'
        )
        self.tree.heading('Data', text='Data')
        self.tree.heading('Horario', text='Horário')
        self.tree.heading('Preco Final', text='Preço Final')
        self.tree.heading('Metodo de pagamento', text='Método de Pagamento')

        # Hide the 'Produtos' column
        self.tree['displaycolumns'] = ('Data', 'Horario', 'Preco Final', 'Metodo de pagamento')

        # Set column widths
        self.tree.column('Data', width=100, anchor=tk.CENTER)
        self.tree.column('Horario', width=100, anchor=tk.CENTER)
        self.tree.column('Preco Final', width=100, anchor=tk.CENTER)
        self.tree.column('Metodo de pagamento', width=150, anchor=tk.CENTER)

        # Add vertical scrollbar
        self.v_scroll = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.v_scroll.set)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.pack(fill=tk.BOTH, expand=True)

        # Bind hover event to show products
        self.tooltip = ToolTip(self.tree)

        # Load and display sales history
        self.load_sales_history()

    def load_sales_history(self):
        try:
            sales_history = pd.read_excel('Files/Historico_vendas.xlsx')
            # Sort by 'Data' and 'Horario' in descending order to show the latest sales first
            sales_history = sales_history.sort_values(by=['Data', 'Horario'], ascending=[False, False])

            # Replace 'nan' in 'Metodo de pagamento' with an empty string
            sales_history['Metodo de pagamento'] = sales_history['Metodo de pagamento'].replace({np.nan: ""})

            for _, row in sales_history.iterrows():
                # Format the 'Preco Final' with R$ and two decimal places
                preco_final = row['Preco Final']
                if isinstance(preco_final, (int, float)):
                    preco_final = f"R${preco_final:.2f}"
                else:
                    preco_final = "R$0.00"

                self.tree.insert('', 'end', values=(
                    row['Data'],
                    row['Horario'],
                    preco_final,
                    row['Metodo de pagamento'],
                    row['Produtos']
                ))
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de histórico de vendas não encontrado!")
