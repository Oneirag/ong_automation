import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(0)

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pywinauto
from pywinauto import Application, Desktop
from PIL import Image, ImageTk
import io
import threading
import time
from datetime import datetime

class PyWinAutoAssistant:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Asistente de Automatización PyWinAuto")
        self.root.geometry("1200x800")
        self.root.state('zoomed')  # Add this line to start maximized
        
        self.selected_window = None
        self.screenshot_image = None
        self.recorded_actions = []
        self.elements_cache = {}
        self.selected_rectangle = None  # Add this line to store the current selection rectangle
        
        self.setup_ui()
        
    def setup_ui(self):
        # Crear notebook para pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Pestaña 1: Seleccionar ventana
        self.window_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.window_frame, text="Seleccionar Ventana")
        self.setup_window_selection()
        
        # Pestaña 2: Captura e interacción
        self.capture_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.capture_frame, text="Captura e Interacción")
        self.setup_capture_interface()
        
        # Pestaña 3: Código generado
        self.code_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.code_frame, text="Código Generado")
        self.setup_code_interface()
        
    def setup_window_selection(self):
        # Frame superior con botones
        top_frame = ttk.Frame(self.window_frame)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(top_frame, text="Actualizar Lista", 
                  command=self.refresh_windows).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="Seleccionar Ventana", 
                  command=self.select_window).pack(side=tk.LEFT, padx=5)
        
        # Lista de ventanas
        self.window_listbox = tk.Listbox(self.window_frame, height=20)
        self.window_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar para la lista
        scrollbar = ttk.Scrollbar(self.window_frame, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.window_listbox.yview)
        
        # Etiqueta de estado
        self.status_label = ttk.Label(self.window_frame, text="Haz clic en 'Actualizar Lista' para ver las ventanas")
        self.status_label.pack(pady=5)
        
        # Cargar ventanas al inicio
        self.refresh_windows()
        
    def setup_capture_interface(self):
        # Frame izquierdo para la captura
        left_frame = ttk.Frame(self.capture_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Botones de control
        control_frame = ttk.Frame(left_frame)
        control_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(control_frame, text="Capturar Pantalla", 
                  command=self.capture_screen).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Identificar Elementos", 
                  command=self.identify_elements).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Limpiar Acciones", 
                  command=self.clear_actions).pack(side=tk.LEFT, padx=5)
        
        # Canvas para la captura
        self.canvas = tk.Canvas(left_frame, bg='white', width=600, height=400)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", lambda e: self.handle_canvas_click(e, "left"))
        self.canvas.bind("<Button-3>", lambda e: self.handle_canvas_click(e, "right"))
        
        # Frame derecho para controles
        right_frame = ttk.Frame(self.capture_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        # Información del elemento seleccionado
        ttk.Label(right_frame, text="Elemento Seleccionado:").pack(anchor=tk.W)
        self.element_info = scrolledtext.ScrolledText(right_frame, height=8, width=40)
        self.element_info.pack(fill=tk.X, pady=5)
        
        # Separador antes de las acciones
        ttk.Separator(right_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Frame para clicks
        click_group = ttk.LabelFrame(right_frame, text="Clicks")
        click_group.pack(fill=tk.X, pady=5, padx=5)
        ttk.Button(click_group, text="Click Simple", 
                  command=lambda: self.add_click_action(double=False)).pack(side=tk.LEFT, padx=2, pady=5)
        ttk.Button(click_group, text="Doble Click", 
                  command=lambda: self.add_click_action(double=True)).pack(side=tk.LEFT, padx=2, pady=5)
        
        # Frame para escribir texto
        text_group = ttk.LabelFrame(right_frame, text="Escribir Texto")
        text_group.pack(fill=tk.X, pady=5, padx=5)
        ttk.Label(text_group, text="Texto:").pack(side=tk.LEFT, padx=5, pady=5)
        self.text_entry = ttk.Entry(text_group, width=20)
        self.text_entry.pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(text_group, text="Escribir", 
                  command=self.add_text_action).pack(side=tk.LEFT, padx=5, pady=5)
        
        # Frame para seleccionar valor
        select_group = ttk.LabelFrame(right_frame, text="Seleccionar Valor")
        select_group.pack(fill=tk.X, pady=5, padx=5)
        ttk.Label(select_group, text="Valor:").pack(side=tk.LEFT, padx=5, pady=5)
        self.select_entry = ttk.Combobox(select_group, width=20)
        self.select_entry.pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(select_group, text="Seleccionar", 
                  command=self.add_select_action).pack(side=tk.LEFT, padx=5, pady=5)
        
        # Frame para pausa
        pause_group = ttk.LabelFrame(right_frame, text="Añadir Pausa")
        pause_group.pack(fill=tk.X, pady=5, padx=5)
        ttk.Label(pause_group, text="Segundos:").pack(side=tk.LEFT, padx=5, pady=5)
        self.pause_entry = ttk.Spinbox(pause_group, width=10, from_=0.1, to=300, increment=0.1)
        self.pause_entry.set("1.0")
        self.pause_entry.pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(pause_group, text="Añadir Pausa", 
                  command=self.add_pause_action).pack(side=tk.LEFT, padx=5, pady=5)
        
        # Separador antes de la lista de acciones
        ttk.Separator(right_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Lista de acciones grabadas
        ttk.Label(right_frame, text="Acciones Grabadas:").pack(anchor=tk.W)
        self.actions_listbox = tk.Listbox(right_frame, height=10)
        self.actions_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Separador antes de los botones de control
        ttk.Separator(right_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Frame para los botones de control de acciones
        actions_control_frame = ttk.Frame(right_frame)
        actions_control_frame.pack(fill=tk.X, pady=5)
        ttk.Button(actions_control_frame, text="Ejecutar Acción", 
                  command=self.execute_selected_action).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        ttk.Button(actions_control_frame, text="Borrar Acción", 
                  command=self.delete_selected_action).pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)

    def _select_last_action(self):
        """Selecciona la última acción añadida a la lista"""
        last_index = self.actions_listbox.size() - 1
        if last_index >= 0:
            self.actions_listbox.select_clear(0, tk.END)
            self.actions_listbox.select_set(last_index)
            self.actions_listbox.see(last_index)
            self.actions_listbox.focus_set()


    def execute_selected_action(self):
        """Ejecuta la acción seleccionada en la lista"""
        selection = self.actions_listbox.curselection()
        if not selection:
            messagebox.showwarning("Advertencia", "Selecciona una acción para ejecutar")
            return
            
        try:
            action = self.recorded_actions[selection[0]]
            
            if action['type'] == 'pause':
                time.sleep(action['seconds'])
            else:
                element = action['element']
                
                # Ejecutar la acción según su tipo
                if action['type'] == 'click':
                    element.click_input()
                elif action['type'] == 'double_click':
                    element.click_input(double=True)
                elif action['type'] == 'type_text':
                    element.set_focus().type_keys(action['text'], with_spaces=True)
                elif action['type'] == 'select':
                    self.select_combobox_item(element, action['value'])
                    
                time.sleep(0.5)  # Esperar a que se complete la acción
                
                # Tomar nueva captura de pantalla
                self.capture_screen()
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al ejecutar la acción: {str(e)}")

    def delete_selected_action(self):
        """Borra la acción seleccionada de la lista"""
        selection = self.actions_listbox.curselection()
        if not selection:
            messagebox.showwarning("Advertencia", "Selecciona una acción para borrar")
            return
            
        # Borrar la acción de la lista y del registro
        self.actions_listbox.delete(selection[0])
        self.recorded_actions.pop(selection[0])
        self._select_last_action()  # Seleccionar la última acción después de borrar

        
    def setup_code_interface(self):
        # Botones superiores
        button_frame = ttk.Frame(self.code_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Generar Código", 
                  command=self.generate_code).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Guardar Código", 
                  command=self.save_code).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Ejecutar Código", 
                  command=self.execute_code).pack(side=tk.LEFT, padx=5)
        
        # Área de código
        self.code_text = scrolledtext.ScrolledText(self.code_frame, height=30, width=100)
        self.code_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def refresh_windows(self):
        """Actualiza la lista de ventanas abiertas"""
        self.window_listbox.delete(0, tk.END)
        
        try:
            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            
            for i, window in enumerate(windows):
                # Ignore minimized windows (width or height is 0)
                if window.rectangle().width() == 0 or window.rectangle().height() == 0:
                    continue
                try:
                    title = window.window_text()
                    if title and title.strip():
                        self.window_listbox.insert(tk.END, f"{i}: {title}")
                except:
                    continue
                    
            self.status_label.config(text=f"Encontradas {self.window_listbox.size()} ventanas")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al obtener ventanas: {str(e)}")
            
    def select_window(self):
        """Selecciona la ventana elegida"""
        selection = self.window_listbox.curselection()
        if not selection:
            messagebox.showwarning("Advertencia", "Selecciona una ventana de la lista")
            return
            
        try:
            window_text = self.window_listbox.get(selection[0])
            window_index = int(window_text.split(":")[0])
            window_full_text = window_text.split(": ")[1]

            desktop = Desktop(backend="uia")
            windows = desktop.windows()
            found_windows = list(w for w in windows if w.window_text() == window_full_text)
            self.selected_window = found_windows[0]
            
            # Traer la ventana al frente
            self.selected_window.set_focus()
            
            self.status_label.config(text=f"Ventana seleccionada: {self.selected_window.window_text()}")
            
            # Cambiar a la pestaña de captura
            self.notebook.select(1)
            self.root.lift()    # Devolver el foco a la ventana principal
            self.root.focus_force()        # Fuerza el foco en la ventana
            # time.sleep(0.5)  # Pequeña pausa para asegurar que la ventana está activa
            # self.capture_screen()  # Capturar pantalla de la ventana seleccionada

        except Exception as e:
            messagebox.showerror("Error", f"Error al seleccionar ventana: {str(e)}")
            
    def capture_screen(self):
        """Captura la pantalla de la ventana seleccionada"""
        if not self.selected_window:
            messagebox.showwarning("Advertencia", "Primero selecciona una ventana")
            return
        
        try:
            # Remove selection rectangle if exists
            if self.selected_rectangle:
                self.canvas.delete(self.selected_rectangle)
                self.selected_rectangle = None
            
            # Hacer la ventana transparente
            original_attributes = self.root.attributes()
            self.root.attributes('-alpha', 0.0)  # Make completely transparent
            
            # Capturar la pantalla
            time.sleep(0.2)  # Pequeña pausa para que se aplique la transparencia
            screenshot = self.selected_window.capture_as_image()
            
            # Restaurar la ventana
            self.root.attributes('-alpha', 1.0)  # Restore opacity
            # Redimensionar para el canvas
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width > 1 and canvas_height > 1:
                screenshot = screenshot.resize((canvas_width, canvas_height), Image.Resampling.LANCZOS)
            
            # Convertir a PhotoImage
            self.screenshot_image = ImageTk.PhotoImage(screenshot)
            
            # Mostrar en el canvas
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.screenshot_image)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al capturar pantalla: {str(e)}")
            

    def identify_elements(self):
        """Identifica elementos en la ventana seleccionada"""
        if not self.selected_window:
            messagebox.showwarning("Advertencia", "Primero selecciona una ventana")
            return
            
        try:
            # Limpiar elementos anteriores
            self.elements_cache.clear()
            
            # Obtener todos los elementos
            elements = self.selected_window.iter_descendants()
            # Tipos de control interactivos
            interactive_types = [
                "Button", "Edit", "ComboBox", "List", "ListItem", 
                "CheckBox", "RadioButton", "Slider", "Spinner",
                "Tab", "TreeItem", # "Custom", "Group"
            ]
            
             # Obtener todos los elementos
            elements_buffer = []
            window_rect = self.selected_window.rectangle()

            for i, buffered_element in enumerate(elements):
                if not buffered_element.is_visible():
                    continue
                if not buffered_element.rectangle().width() > 0:
                    continue
                elements_buffer.append(buffered_element)
                # Process elements in batches of 10
                if len(elements_buffer) >= 20:
                    for element in elements_buffer:
                        try:
                            rect = element.rectangle()
                            
                            # Calcular posición relativa
                            rel_x = rect.left - window_rect.left
                            rel_y = rect.top - window_rect.top
                            
                            # Escalar a coordenadas del canvas
                            canvas_width = self.canvas.winfo_width()
                            canvas_height = self.canvas.winfo_height()
                            
                            if canvas_width > 1 and canvas_height > 1:
                                scale_x = canvas_width / window_rect.width()
                                scale_y = canvas_height / window_rect.height()
                                
                                canvas_x = rel_x * scale_x
                                canvas_y = rel_y * scale_y

                                is_interactive = element.element_info.control_type in interactive_types
                                # Crear rectángulo en el canvas
                                rect_id = self.canvas.create_rectangle(
                                    canvas_x, canvas_y, 
                                    canvas_x + rect.width() * scale_x, 
                                    canvas_y + rect.height() * scale_y,
                                    outline='red' if is_interactive else "gold",  # Semi-transparent blue
                                    width=2 if is_interactive else 0.5,
                                    dash=(1,1) if not is_interactive else None  # Dotted line for non-interactive
                                )

                                if is_interactive:                                          
                                    # Añadir texto con el tipo de control
                                    self.canvas.create_text(
                                        canvas_x + 5,
                                        canvas_y + 5,
                                        text=element.element_info.control_type,
                                        anchor=tk.NW,
                                        fill='red',
                                        font=('Arial', 8)
                                    )

                                
                                # Crear punto en el canvas
                                point_id = self.canvas.create_oval(canvas_x-3, canvas_y-3, 
                                                                canvas_x+3, canvas_y+3, 
                                                                fill='red', outline='darkred')
                                
                                # Guardar información del elemento
                                self.elements_cache[point_id] = {
                                    'element': element,
                                    'canvas_pos': (canvas_x, canvas_y),
                                    'window_pos': (rel_x, rel_y)
                                }
                                    
                        except Exception as e:
                            print(f"Error procesando elemento: {str(e)}")
                            continue
                        # Update the canvas and clear the buffer
                        self.canvas.update()
                        elements_buffer = []
                    
            messagebox.showinfo("Info", f"Identificados {len(self.elements_cache)} elementos")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al identificar elementos: {str(e)}")

            
    def handle_canvas_click(self, event, button_type):
        """Common handler for both left and right clicks on canvas"""
        try:
            if button_type == "left":
                # Find closest element using canvas coordinates
                clicked_item = self.canvas.find_closest(event.x, event.y)[0]
                if clicked_item not in self.elements_cache:
                    return
                element_info = self.elements_cache[clicked_item]
                element = element_info['element']
            else:  # right click
                # Convert canvas coordinates to window coordinates
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()
                window_rect = self.selected_window.rectangle()
                
                scale_x = window_rect.width() / canvas_width
                scale_y = window_rect.height() / canvas_height
                
                window_x = int(event.x * scale_x) + window_rect.left
                window_y = int(event.y * scale_y) + window_rect.top
                
                # Get element at point
                self.selected_window.set_focus()
                time.sleep(0.1)
                element = self.selected_window.from_point(window_x, window_y)
                self.root.lift()
                self.root.focus_force()
                
                if not element:
                    return
                    
                element_info = {
                    'element': element,
                    'window_pos': (window_x - window_rect.left, window_y - window_rect.top),
                    'canvas_pos': (event.x, event.y)
                }

            # Common code for handling the selected element
            self.process_selected_element(element, element_info)
            
        except Exception as e:
            print(f"Error handling canvas click: {e}")

    def process_selected_element(self, element, element_info):
        """Process the selected element regardless of how it was selected"""
        
        # Remove previous selection rectangle if exists
        if self.selected_rectangle:
            self.canvas.delete(self.selected_rectangle)
    
        # Get element rectangle and convert to canvas coordinates
        rect = element.rectangle()
        window_rect = self.selected_window.rectangle()
    
        # Calculate relative position and size
        rel_x = rect.left - window_rect.left
        rel_y = rect.top - window_rect.top
        rel_width = rect.width()
        rel_height = rect.height()
    
        # Convert to canvas coordinates
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
    
        if canvas_width > 1 and canvas_height > 1:
            scale_x = canvas_width / window_rect.width()
            scale_y = canvas_height / window_rect.height()
        
            canvas_x = rel_x * scale_x
            canvas_y = rel_y * scale_y
            canvas_width = rel_width * scale_x
            canvas_height = rel_height * scale_y
        
            # Create selection rectangle
            self.selected_rectangle = self.canvas.create_rectangle(
                canvas_x, canvas_y,
                canvas_x + canvas_width,
                canvas_y + canvas_height,
                outline='green',
                width=2
            )
        
            # Ensure rectangle is on top
            self.canvas.tag_raise(self.selected_rectangle)
    
        # Display element information
        info_text = f"Tipo: {element.element_info.control_type}\n"
        info_text += f"Nombre: {element.window_text()}\n"
        info_text += f"Clase: {element.element_info.class_name}\n"
        info_text += f"ID: {element.element_info.automation_id}\n"
        info_text += f"Handle: {element.element_info.handle}\n"
        info_text += f"Posición: {element_info['window_pos']}\n"
        
        self.element_info.delete(1.0, tk.END)
        self.element_info.insert(tk.END, info_text)
        
        # Handle ComboBox elements
        if element.element_info.control_type == "ComboBox":
            try:
                # Try multiple approaches to get combobox items
                items = []
                
                # Try getting items from select patterns
                try:
                    select_pattern = element.iface_selection
                    if select_pattern:
                        items = [opt.name for opt in select_pattern.GetSelection()]
                except:
                    try:
                        # Expand combobox and find specific item
                        element.expand()
                        time.sleep(0.5) # Wait for items to be visible
                        list_items = element.descendants(control_type="ListItem")
                        items = [item.window_text() for item in list_items]
                        element.collapse()
                    except:
                        # As last resort, try getting items from parent 
                        try:
                            parent = element.parent()
                            list_items = parent.descendants(control_type="ListItem") 
                            items = [item.window_text() for item in list_items]
                        except:
                            print("Could not get ComboBox items")
                            items = []
                self.select_entry['values'] = items
                self.select_entry.state(['!disabled'])
            except Exception as e:
                print(f"Error getting ComboBox items: {e}")
                self.select_entry.set('')
                self.select_entry['values'] = []
                self.select_entry.state(['disabled'])
        else:
            self.select_entry.set('')
            self.select_entry['values'] = []
            self.select_entry.state(['disabled'])
        
        # Store selected element
        self.current_element = element_info
    
    def select_combobox_item(self, element, desired_text):
        """Ayuda a seleccionar un item específico de un ComboBox"""
        try:
            element.expand()
            time.sleep(0.5)  # Esperar a que los items sean visibles
            list_items = element.descendants(control_type="ListItem")
            for item in list_items:
                if item.window_text() == desired_text:
                    item.click_input()
                    break
            else:  # Si no se encontró el item
                element.collapse()
        except Exception as e:
            print(f"Error al seleccionar item del ComboBox: {e}")

    def add_click_action(self, double=False):
        """Añade una acción de click o doble click"""
        if not hasattr(self, 'current_element'):
            messagebox.showwarning("Advertencia", "Primero selecciona un elemento")
            return
            
                # Verificar si el elemento es interactivo
        interactive_types = [
            "Button", "Edit", "ComboBox", "List", "ListItem", 
            "CheckBox", "RadioButton", "Slider", "Spinner",
            "Tab", "TreeItem", "Hyperlink"
        ]
        
        is_interactive = self.current_element['element'].element_info.control_type in interactive_types
        
        if not is_interactive:
            if not (response := messagebox.askyesno(
                "Elemento no interactivo",
                "Este elemento no parece ser interactivo. ¿Deseas añadir la acción de todos modos?"
            )):
                return

        action = {
            'type': 'double_click' if double else 'click',
            'element': self.current_element['element'],
            'position': self.current_element['window_pos']
        }
        
        self.recorded_actions.append(action)
        self.actions_listbox.insert(tk.END, 
            f"Doble click en {action['element'].window_text()}" if double 
            else f"Click en {action['element'].window_text()}")
        self._select_last_action()  # Añade esta línea

    def add_text_action(self):
        """Añade una acción de escribir texto"""
        if not hasattr(self, 'current_element'):
            messagebox.showwarning("Advertencia", "Primero selecciona un elemento")
            return
            
        text = self.text_entry.get()
        if not text:
            messagebox.showwarning("Advertencia", "Ingresa el texto a escribir")
            return
            
        action = {
            'type': 'type_text',
            'element': self.current_element['element'],
            'text': text
        }
        
        self.recorded_actions.append(action)
        self.actions_listbox.insert(tk.END, f"Escribir: '{text}'")
        self.text_entry.delete(0, tk.END)
        self._select_last_action()  # Añade esta línea

    def add_select_action(self):
        """Añade una acción de seleccionar valor"""
        if not hasattr(self, 'current_element'):
            messagebox.showwarning("Advertencia", "Primero selecciona un elemento")
            return
            
        value = self.select_entry.get()
        if not value:
            messagebox.showwarning("Advertencia", "Selecciona un valor de la lista")
            return
            
        action = {
            'type': 'select',
            'element': self.current_element['element'],
            'value': value
        }
        
        self.recorded_actions.append(action)
        self.actions_listbox.insert(tk.END, f"Seleccionar: '{value}'")
        self.select_entry.set('')
        self._select_last_action()  # Añade esta línea

    def add_pause_action(self):
        """Añade una acción de pausa"""
        try:
            seconds = float(self.pause_entry.get())
            if seconds <= 0:
                messagebox.showwarning("Advertencia", "La duración debe ser mayor que 0")
                return
                
            action = {
                'type': 'pause',
                'seconds': seconds
            }
            
            self.recorded_actions.append(action)
            self.actions_listbox.insert(tk.END, f"Pausa: {seconds} segundos")
            self._select_last_action()  # Añade esta línea
            
        except ValueError:
            messagebox.showwarning("Advertencia", "Ingresa un número válido de segundos")

    def clear_actions(self):
        """Limpia las acciones grabadas"""
        self.recorded_actions.clear()
        self.actions_listbox.delete(0, tk.END)
    
    def pywinauto_selected_window(self) -> str:
        """Devuelve un codigo de python para identificar la ventana seleccionada"""
        try:
            window = Desktop(backend='uia').window(title=self.selected_window.window_text())
            return f"window = desktop.window(title='{self.selected_window.window_text()}')"
        except Exception as e:
            # No he podido identificar la ventana, probar de otra forma
            window = self.recorded_actions[0]['element'].top_level_window()
            return f"window = desktop.window(handle={window.element_info.handle})"
        
    def pywinauto_element_selector(self, idx: int, element) -> str:
        """Devuelve un código de python para identificar un elemento"""
        automation_id = element.element_info.automation_id
        automation_id = None    # UIA does not support automation_id as a selector
        class_name = element.element_info.class_name
        control_type = element.element_info.control_type
        window_text = element.window_text() 
        selectors = {}
        
        if automation_id:
            selectors["auto_id"] = automation_id
        if class_name:
            selectors['class_name'] = class_name
        if control_type:
            selectors["control_type"] = control_type
        if window_text:
            selectors["title"] = window_text  
        

        if element.is_visible():
            # If selector provides more than one element, we need to add an index
            selected_elements = self.selected_window.descendants(**selectors)
            if len(selected_elements) > 1:   
                # Iterate all children to find the correct one 
                for i, child in enumerate(selected_elements):
                    # check for the same position
                        if child.rectangle() == element.rectangle():
                            idx = i
                            break
                        selectors['found_index'] = idx

        def to_str(v) -> str:
            if isinstance(v, int):
                return v
            else:
                return f"'{v}'"
            
        selector = ", ".join(f"{k}={to_str(v)}" for k, v in selectors.items())        
        return f"element_{idx} = window.child_window({selector})"
        
    
    def generate_code(self):
        """Genera el código PyWinAuto"""
        if not self.recorded_actions:
            messagebox.showwarning("Advertencia", "No hay acciones grabadas")
            return
            
        if not self.selected_window:
            messagebox.showwarning("Advertencia", "No hay ventana seleccionada")
            return
        if not self.selected_window.is_visible():
            messagebox.showwarning("Advertencia", "La ventana seleccionada no está visible. "
                                       "No se puede generar código para ventanas no visibles")
            return
            
        # Generar código
        code = "from pywinauto import Application, Desktop\n"
        code += "import time\n\n"

        code += "# This global parameter is required to execute the code using eval"
        code += "global select_combobox_item\n"
        code += "def select_combobox_item(element, desired_text):\n"
        code += "    try:\n"
        code += "        element.expand()\n"
        code += "        time.sleep(0.5) # Wait for items to be visible\n"
        code += "        list_items = element.descendants(control_type=\"ListItem\")\n"
        code += "        for item in list_items:\n"
        code += "            if item.window_text() == desired_text:\n"
        code += "                item.click_input()\n"
        code += "                break\n"
        code += "        # If we didn't find it, collapse the combobox\n"
        code += "        else:\n"
        code += "            element.collapse()\n"
        code += "    except Exception as e:\n"
        code += "        print(f\"Error al seleccionar item del ComboBox: {e}\")\n"
        
        code += "\ndef automatizar_ventana():\n"
        code += "    # Won't work without global in eval\n"
        code += "    global select_combobox_item\n"
        code += "    try:\n"
        code += f"        # Conectar a la ventana: {self.selected_window.window_text()}\n"
        code += "        desktop = Desktop(backend='uia')\n"
        code += f"        {self.pywinauto_selected_window()}\n"
        code += "        window.set_focus()\n"
        code += "        time.sleep(1)\n\n"
        
        for i, action in enumerate(self.recorded_actions):
            code += f"        # Acción {i+1}: {action['type']}\n"
            
            if action['type'] == 'pause':
                code += f"        time.sleep({action['seconds']})  # Pausa explícita\n"
            else:
                # Obtener identificadores del elemento
                element = action['element']
                if not element.is_visible():
                    messagebox.showwarning("Advertencia", f"El elemento {i+1} no es visible. "
                                       "El código generado podría no funcionar adecuadamente")
                code += f"        {self.pywinauto_element_selector(i, element)}\n"
                
                if action['type'] == 'click':
                    code += f"        element_{i}.wrapper_object().click_input()\n"
                elif action['type'] == 'double_click':
                    code += f"        element_{i}.wrapper_object().click_input(double=True)\n"
                elif action['type'] == 'type_text':
                    code += f"        element_{i}.set_focus().type_keys('{action['text']}', with_spaces=True)\n"
                elif action['type'] == 'select':
                    code += f"        select_combobox_item(element_{i}, '{action['value']}')\n"
                    
            code += "        time.sleep(0.5)\n\n"
            
        code += "        print('Automatización completada exitosamente')\n"
        code += "    except Exception as e:\n"
        code += "        print(f'Error en la automatización: {e}')\n\n"
        
        code += "if __name__ == '__main__':\n"
        
        code += "    automatizar_ventana()\n"
        
        # Mostrar código
        self.code_text.delete(1.0, tk.END)
        self.code_text.insert(tk.END, code)
        
        # Cambiar a la pestaña de código
        self.notebook.select(2)
        
    def save_code(self):
        """Guarda el código en un archivo"""
        code = self.code_text.get(1.0, tk.END)
        
        if not code.strip():
            messagebox.showwarning("Advertencia", "No hay código para guardar")
            return
            
        from tkinter import filedialog
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(code)
                messagebox.showinfo("Éxito", f"Código guardado en: {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar: {str(e)}")
                
    def execute_code(self):
        """Ejecuta el código generado"""
        code = self.code_text.get(1.0, tk.END)
        
        if not code.strip():
            messagebox.showwarning("Advertencia", "No hay código para ejecutar")
            return
            
        try:
            res = exec(code)
            messagebox.showinfo("Éxito", f"Código ejecutado correctamente:\n{res}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al ejecutar: {str(e)}")
            
    def run(self):
        """Ejecuta la aplicación"""
        self.root.mainloop()

# Ejecutar la aplicación
if __name__ == "__main__":
    app = PyWinAutoAssistant()
    app.run()