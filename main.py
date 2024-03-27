import os
import math
import numpy as np
from itertools import cycle
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter.filedialog as tkfiledialog
import uuid
import material_properties as mp
import openpyxl
from openpyxl import Workbook


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class CompositeMaterialsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Layup")
        self.root.tk.call("source", resource_path("themes\\azure.tcl"))
        self.root.tk.call("set_theme", "dark")

        # Frame setup
        self.setup_frames()

        # Notebook and tab setup
        self.setup_notebooks_and_tabs()

        # Button setup
        self.setup_layup_entry()

        # Layup data
        self.setup_layup_data_tree()

        # Layup failure
        self.setup_layup_failure_tree()

        # Layup modification
        self.setup_layup_modification()

        # Matrix label setup
        self.setup_matrix_labels()

        # Set up on-axis data trees
        self.setup_on_axis_stress_data_tree()
        self.setup_on_axis_strain_data_tree()
        self.setup_on_axis_strength_prop_data_tree()
        self.setup_on_axis_material_prop_data_tree()

        # Set up off-axis data trees
        self.setup_off_axis_strain_data_tree()

        # Set stress and moment
        self.set_stress()
        self.set_moment()

    def setup_frames(self):
        # Top frame
        self.top_frame = ttk.Frame(root)
        self.top_frame.pack(side=tk.TOP, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Top Left Sub-Frame
        self.top_left_frame = ttk.Frame(self.top_frame)
        self.top_left_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Top Right Sub-Frame
        self.top_right_frame = ttk.Frame(self.top_frame)
        self.top_right_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Bottom frame
        self.bottom_frame = ttk.Frame(root)
        self.bottom_frame.pack(side=tk.BOTTOM, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Bottom Left Sub-Frame
        self.bottom_left_frame = ttk.Frame(self.bottom_frame)
        self.bottom_left_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Bottom Left Sub-Sub-Frame
        self.bottom_bottom_left_frame = ttk.Frame(self.bottom_left_frame)
        self.bottom_bottom_left_frame.pack(side=tk.BOTTOM, padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Bottom Right Sub-Frame
        self.bottom_right_frame = ttk.Frame(self.bottom_frame)
        self.bottom_right_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

    def setup_notebooks_and_tabs(self):
        # User input
        self.user_input_label = ttk.Label(self.top_left_frame, text="User input", font="Helvetica 14 bold")
        self.user_input_label.pack(side=tk.TOP, anchor="nw", padx=5, pady=10)

        self.user_input_notebook = ttk.Notebook(self.top_left_frame)
        self.user_input_notebook.pack(fill=tk.BOTH, expand=True)

        self.layup_creation_tab = ttk.Frame(self.user_input_notebook)
        self.user_input_notebook.add(self.layup_creation_tab, text="Layup creator")

        self.stress_state_tab = ttk.Frame(self.user_input_notebook)
        self.user_input_notebook.add(self.stress_state_tab, text="Set stress")

        self.layup_options_tab = ttk.Frame(self.user_input_notebook)
        self.user_input_notebook.add(self.layup_options_tab, text="Layup options")

        # Layup
        self.layup_label = ttk.Label(self.bottom_left_frame, text="Layup", font="Helvetica 14 bold")
        self.layup_label.pack(side=tk.TOP, anchor="nw", padx=5, pady=10)

        self.layup_notebook = ttk.Notebook(self.bottom_left_frame)
        self.layup_notebook.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.layup_tab = ttk.Frame(self.layup_notebook)
        self.layup_notebook.add(self.layup_tab, text="Layer count: 0")

        self.layup_failure_tab = ttk.Frame(self.layup_notebook)
        self.layup_notebook.add(self.layup_failure_tab, text="Layup failure")

        # On-axis properties
        self.on_axis_label = ttk.Label(self.top_right_frame, text="On-axis and material properties",
                                       font="Helvetica 14 bold")
        self.on_axis_label.pack(side=tk.TOP, anchor="nw", padx=5, pady=10)

        self.on_axis_notebook = ttk.Notebook(self.top_right_frame)
        self.on_axis_notebook.pack(fill=tk.BOTH, expand=True)

        self.on_axis_Q_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_Q_tab, text="Q")

        self.on_axis_S_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_S_tab, text="S")

        self.on_axis_stress_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_stress_tab, text="Stress")

        self.on_axis_strain_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_strain_tab, text="Strain")

        self.on_axis_strength_prop_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_strength_prop_tab, text="Strength")

        self.on_axis_material_prop_tab = ttk.Frame(self.on_axis_notebook)
        self.on_axis_notebook.add(self.on_axis_material_prop_tab, text="Material")

        # Off-axis properties
        self.off_axis_label = ttk.Label(self.bottom_right_frame, text="Off-axis properties", font="Helvetica 14 bold")
        self.off_axis_label.pack(side=tk.TOP, anchor="nw", padx=5, pady=10)

        self.off_axis_notebook = ttk.Notebook(self.bottom_right_frame)
        self.off_axis_notebook.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

        self.off_axis_Q_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_Q_tab, text="Q")

        self.off_axis_S_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_S_tab, text="S")

        self.off_axis_A_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_A_tab, text="A")

        self.off_axis_a_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_a_tab, text="a")

        self.off_axis_D_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_D_tab, text="D")

        self.off_axis_d_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_d_tab, text="d")

        self.off_axis_strain_tab = ttk.Frame(self.off_axis_notebook)
        self.off_axis_notebook.add(self.off_axis_strain_tab, text="Strain")

    def setup_layup_entry(self):
        # Material drop down menu
        self.material_label = ttk.Label(self.layup_creation_tab, text="Material:")
        self.material_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.material_var = tk.StringVar()
        self.material_dropdown = ttk.Combobox(self.layup_creation_tab, textvariable=self.material_var,
                                              values=["T300/5208", "B4/5505", "AS/H3501", "Scotchply 1002",
                                                      "Kevlar49/epoxy"], state="readonly")
        self.material_dropdown.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Ply thickness
        self.thickness_label = ttk.Label(self.layup_creation_tab, text="Thickness (mm):")
        self.thickness_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.thickness_var = tk.StringVar(value="4.925")  # Set default value to 4.925
        self.thickness_entry = ttk.Entry(self.layup_creation_tab, textvariable=self.thickness_var, justify="center")
        self.thickness_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        # Ply orientation
        self.orientation_label = ttk.Label(self.layup_creation_tab, text="Orientation (deg):")
        self.orientation_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.orientation_var = tk.StringVar()
        self.orientation_entry = ttk.Entry(self.layup_creation_tab, textvariable=self.orientation_var, justify="center")
        self.orientation_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        # Enter ply
        self.enter_button = ttk.Button(self.layup_creation_tab, text="Enter", command=self.add_to_layup)
        self.enter_button.grid(row=6, column=2, pady=5, padx=5, sticky="e")

        # Has core label
        self.has_core_label = ttk.Label(self.layup_options_tab, text="Has Core:")
        self.has_core_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Has core checkbox
        self.has_core_var = tk.BooleanVar()
        self.has_core_checkbox = ttk.Checkbutton(self.layup_options_tab, text="", variable=self.has_core_var,
                                                 command=self.toggle_core)
        self.has_core_checkbox.grid(row=0, column=1, pady=5, sticky="w")

        # Core thickness
        self.core_thickness_label = ttk.Label(self.layup_options_tab, text="Core Thickness (mm):")
        self.core_thickness_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.core_thickness_var = tk.StringVar()
        self.core_thickness_entry = ttk.Entry(self.layup_options_tab, textvariable=self.core_thickness_var,
                                              justify="center", state="disabled")
        self.core_thickness_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Stress state
        self.N_1_label = ttk.Label(self.stress_state_tab, text="N₁ (N/m)")
        self.N_1_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.N_1_var = tk.StringVar(value="0")
        self.N_1_entry = ttk.Entry(self.stress_state_tab, textvariable=self.N_1_var, justify="center",
                                       width=8)
        self.N_1_entry.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        self.N_2_label = ttk.Label(self.stress_state_tab, text="N₂ (N/m)")
        self.N_2_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.N_2_var = tk.StringVar(value="0")
        self.N_2_entry = ttk.Entry(self.stress_state_tab, textvariable=self.N_2_var, justify="center",
                                       width=8)
        self.N_2_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        self.N_6_label = ttk.Label(self.stress_state_tab, text="N₆ (N/m)")
        self.N_6_label.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.N_6_var = tk.StringVar(value="0")
        self.N_6_entry = ttk.Entry(self.stress_state_tab, textvariable=self.N_6_var, justify="center",
                                       width=8)
        self.N_6_entry.grid(row=1, column=2, padx=5, pady=5, sticky="w")

        # Moment
        self.M_1_label = ttk.Label(self.stress_state_tab, text="M₁ (N)")
        self.M_1_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.M_1_var = tk.StringVar(value="0")
        self.M_1_entry = ttk.Entry(self.stress_state_tab, textvariable=self.M_1_var, justify="center",
                                       width=8)
        self.M_1_entry.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        self.M_2_label = ttk.Label(self.stress_state_tab, text="M₂ (N)")
        self.M_2_label.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        self.M_2_var = tk.StringVar(value="0")
        self.M_2_entry = ttk.Entry(self.stress_state_tab, textvariable=self.M_2_var, justify="center",
                                       width=8)
        self.M_2_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        self.M_6_label = ttk.Label(self.stress_state_tab, text="M₆ (N)")
        self.M_6_label.grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.M_6_var = tk.StringVar(value="0")
        self.M_6_entry = ttk.Entry(self.stress_state_tab, textvariable=self.M_6_var, justify="center",
                                       width=8)
        self.M_6_entry.grid(row=3, column=2, padx=5, pady=5, sticky="w")

        # Submit Buttons
        self.stress_set_button = ttk.Button(self.stress_state_tab, text="Set stress", command=self.set_stress)
        self.stress_set_button.grid(row=1, column=3, pady=5, padx=5, sticky="e")

        self.moment_set_button = ttk.Button(self.stress_state_tab, text="Set moment", command=self.set_moment)
        self.moment_set_button.grid(row=3, column=3, pady=5, padx=5, sticky="e")

    def setup_layup_data_tree(self):
        # Frame to contain the Treeview and Scrollbar
        self.layup_data_frame = ttk.Frame(self.layup_tab)
        self.layup_data_frame.grid(row=6, column=0, padx=5, pady=5, columnspan=3, sticky="nsew")

        # Treeview
        self.data_grid = ttk.Treeview(self.layup_data_frame,
                                      columns=("Ply Number", "Ply Type", "Thickness", "Orientation"),
                                      show="headings")
        self.data_grid.heading("Ply Number", text="#")
        self.data_grid.heading("Ply Type", text="Material")
        self.data_grid.heading("Thickness", text="Thickness (mm)")
        self.data_grid.heading("Orientation", text="Orientation (deg)")
        self.data_grid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Center the text in the Treeview cells
        for column in ("Ply Number", "Ply Type", "Thickness", "Orientation"):
            self.data_grid.column(column, anchor="center")

        # Set fixed width for each column
        self.data_grid.column("Ply Number", width=25)
        self.data_grid.column("Ply Type", width=85)
        self.data_grid.column("Thickness", width=95)
        self.data_grid.column("Orientation", width=115)  # Adjust the width as needed

        # Add a vertical scrollbar
        data_scrollbar = ttk.Scrollbar(self.layup_data_frame, orient="vertical", command=self.data_grid.yview)
        data_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.data_grid.configure(yscrollcommand=data_scrollbar.set)

        # Configure the data_frame to expand both vertically and horizontally
        self.layup_tab.columnconfigure(0, weight=1)
        self.layup_tab.rowconfigure(6, weight=1)

    def setup_layup_failure_tree(self):
        # Frame to contain the Treeview and Scrollbar
        self.layup_failure_frame = ttk.Frame(self.layup_failure_tab)
        self.layup_failure_frame.grid(row=6, column=0, padx=5, pady=5, columnspan=3, sticky="nsew")

        # Treeview
        self.failure_grid = ttk.Treeview(self.layup_failure_frame,
                                      columns=("Ply Number", "Orientation", "F.I.x", "F.I.y", "MOF"),
                                      show="headings")
        self.failure_grid.heading("Ply Number", text="#")
        self.failure_grid.heading("Orientation", text="Orientation (deg)")
        self.failure_grid.heading("F.I.x", text="F.I.x")
        self.failure_grid.heading("F.I.y", text="F.I.y")
        self.failure_grid.heading("MOF", text="MOF")
        self.failure_grid.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Center the text in the Treeview cells
        for column in ("Ply Number", "Orientation", "F.I.x", "F.I.y", "MOF"):
            self.failure_grid.column(column, anchor="center")

        # Set fixed width for each column
        self.failure_grid.column("Ply Number", width=25)
        self.failure_grid.column("Orientation", width=115)  # Adjust the width as needed
        self.failure_grid.column("F.I.x", width=25)  # Adjust the width as needed
        self.failure_grid.column("F.I.y", width=25)  # Adjust the width as needed
        self.failure_grid.column("MOF", width=25)  # Adjust the width as needed

        # Add a vertical scrollbar
        failure_scrollbar = ttk.Scrollbar(self.layup_failure_frame, orient="vertical", command=self.failure_grid.yview)
        failure_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.failure_grid.configure(yscrollcommand=failure_scrollbar.set)

        # Configure the data_frame to expand both vertically and horizontally
        self.layup_failure_tab.columnconfigure(0, weight=1)
        self.layup_failure_tab.rowconfigure(6, weight=1)

    def setup_on_axis_stress_data_tree(self):
        self.on_axis_stress_tree = ttk.Treeview(self.on_axis_stress_tab, columns=("sigma_x", "sigma_y", "sigma_s"), show="headings")

        # Set headings and column options
        headings = ["σₓ (MPa)", "σᵧ (MPa)", "σₛ (MPa)"]
        for col, heading in enumerate(headings):
            self.on_axis_stress_tree.heading(col, text=heading)
            self.on_axis_stress_tree.column(col, width=50, anchor="center")  # Center the data

        # Set a maximum height for the treeview widget
        self.on_axis_stress_tree.config(height=1)  # Adjust the value as needed

        self.on_axis_stress_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.on_axis_stress_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def setup_on_axis_strain_data_tree(self):
        self.on_axis_strain_tree = ttk.Treeview(self.on_axis_strain_tab, columns=("epsilon_x", "epsilon_y", "epsilon_s"),
                                        show="headings")

        # Set headings and column options
        headings = ["εₓ", "εᵧ", "εₛ"]
        for col, heading in enumerate(headings):
            self.on_axis_strain_tree.heading(col, text=heading)
            self.on_axis_strain_tree.column(col, width=50, anchor="center")  # Center the data

        # Set a maximum height for the treeview widget
        self.on_axis_strain_tree.config(height=1)  # Adjust the value as needed

        self.on_axis_strain_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.on_axis_strain_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def setup_on_axis_strength_prop_data_tree(self):
        self.strength_tree = ttk.Treeview(self.on_axis_strength_prop_tab, columns=("X_t", "Y_t", "X_c", "Y_c", "S_c"),
                                          show="headings")

        # Set headings and column options
        headings = ["Xₜ (MPa)", "Yₜ (MPa)", "X꜀ (MPa)", "Y꜀ (MPa)", "S꜀ (MPa)"]
        for col, heading in enumerate(headings):
            self.strength_tree.heading(col, text=heading)
            self.strength_tree.column(col, width=50, anchor="center")  # Center the data

        # Set a maximum height for the treeview widget
        self.strength_tree.config(height=1)  # Adjust the value as needed

        self.strength_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.strength_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def setup_on_axis_material_prop_data_tree(self):
        self.material_properties_tree = ttk.Treeview(self.on_axis_material_prop_tab, columns=("E_x", "E_y", "E_s", "ν"),
                                                     show="headings")

        # Set headings and column options
        headings = ["Eₓ (GPa)", "Eᵧ (GPa)", "Eₛ (GPa)", "ν"]
        for col, heading in enumerate(headings):
            self.material_properties_tree.heading(col, text=heading)
            self.material_properties_tree.column(col, width=62, anchor="center")  # Center the data

        # Set a maximum height for the treeview widget
        self.material_properties_tree.config(height=1)  # Adjust the value as needed

        self.material_properties_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.material_properties_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def setup_off_axis_strain_data_tree(self):
        self.strain_off_tree = ttk.Treeview(self.off_axis_strain_tab, columns=("epsilon_x", "epsilon_y", "epsilon_s"),
                                            show="headings")

        # Set headings and column options
        headings = ["ε₁", "ε₂", "ε₆"]
        for col, heading in enumerate(headings):
            self.strain_off_tree.heading(col, text=heading)
            self.strain_off_tree.column(col, width=50, anchor="center")  # Center the data

        # Set a maximum height for the treeview widget
        self.strain_off_tree.config(height=1)  # Adjust the value as needed

        self.strain_off_tree.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.strain_off_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def setup_layup_modification(self):
        # Arrow button for moving the selected ply up
        self.arrow_up_button = ttk.Button(self.bottom_bottom_left_frame, text="▲", command=self.move_up, width=2)
        self.arrow_up_button.grid(row=0, column=1, pady=5, padx=5, sticky="w")

        # Arrow button for moving the selected ply down
        self.arrow_down_button = ttk.Button(self.bottom_bottom_left_frame, text="▼", command=self.move_down, width=2)
        self.arrow_down_button.grid(row=1, column=1, pady=5, padx=5, sticky="w")

        # Delete button to delete the selected ply
        self.delete_button = ttk.Button(self.bottom_bottom_left_frame, text="Delete", command=self.delete_selected_row)
        self.delete_button.grid(row=1, column=2, pady=5, padx=5, sticky="w")

        # Symmetric button to create a symmetric layup form the current layup
        self.symmetric_button = ttk.Button(self.bottom_bottom_left_frame, text="Symmetric",
                                           command=self.copy_symmetric)
        self.symmetric_button.grid(row=1, column=0, pady=5, padx=5, sticky="w")

        # Delete all button to delete the entire ply
        self.delete_all_button = ttk.Button(self.bottom_bottom_left_frame, text="Delete All",
                                            command=self.delete_all_entries)
        self.delete_all_button.grid(row=1, column=3, pady=5, padx=5, sticky="w")

        # Save button the save the layup to excel
        self.save_button = ttk.Button(self.bottom_bottom_left_frame, text="Save", command=self.save_data)
        self.save_button.grid(row=0, column=3, pady=5, padx=5, sticky="w")

        # Select button
        self.calculate_button = ttk.Button(self.bottom_bottom_left_frame, text="Calculate", command=self.calculate)
        self.calculate_button.grid(row=0, column=0, pady=5, padx=5, sticky="w")

        # Layer top / bottom selector
        self.layer_side_var = tk.StringVar()
        self.layer_side_dropdown = ttk.Combobox(self.bottom_bottom_left_frame, textvariable=self.layer_side_var,
                                              values=["Outer", "Middle", "Inner"], state="readonly")
        self.layer_side_dropdown.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.layer_side_dropdown.current(0)

    def setup_matrix_labels(self):
        # Create and display the matrices
        on_axis_Q_matrix_entries = self.setup_matrix(self.on_axis_Q_tab, "Q", row_labels=["σₓ", "σᵧ", "σₛ"],
                                             column_labels=["εₓ", "εᵧ", "εₛ"],
                                             units=["(GPa)", "(GPa)", "(GPa)"])
        self.on_axis_Q_matrix_entries = on_axis_Q_matrix_entries  # Save q_matrix_entries for later use in calculations

        off_axis_Q_matrix_entries = self.setup_matrix(self.off_axis_Q_tab, "Q", row_labels=["σ₁", "σ₂", "σ₆"],
                                                 column_labels=["ε₁", "ε₂", "ε₆"],
                                                 units=["(GPa)", "(GPa)", "(GPa)"])
        self.off_axis_Q_matrix_entries = off_axis_Q_matrix_entries  # Save q_matrix_entries for later use in calculations

        # Set up S matrix entries
        on_axis_S_matrix_entries = self.setup_matrix(self.on_axis_S_tab, "S", row_labels=["εₓ", "εᵧ", "εₛ"],
                                             column_labels=["σₓ", "σᵧ", "σₛ"],
                                             units=["(GPa⁻¹)", "(GPa⁻¹)", "(GPa⁻¹)"])
        self.on_axis_S_matrix_entries = on_axis_S_matrix_entries

        off_axis_S_matrix_entries = self.setup_matrix(self.off_axis_S_tab, "S", row_labels=["ε₁", "ε₂", "ε₆"],
                                                 column_labels=["σ₁", "σ₂", "σ₆"],
                                                 units=["(GPa⁻¹)", "(GPa⁻¹)", "(GPa⁻¹)"])
        self.off_axis_S_matrix_entries = off_axis_S_matrix_entries

        off_axis_A_matrix_entries = self.setup_matrix(self.off_axis_A_tab, "A", row_labels=["N₁", "N₂", "N₆"],
                                                 column_labels=["ε₁", "ε₂", "ε₆"],
                                                 units=["(GPa·m)", "(GPa·m)", "(GPa·m)"])
        self.off_axis_A_matrix_entries = off_axis_A_matrix_entries  # Save q_matrix_entries for later use in calculations

        off_axis_a_matrix_entries = self.setup_matrix(self.off_axis_a_tab, "a", column_labels=["N₁", "N₂", "N₆"],
                                                 row_labels=["ε₁", "ε₂", "ε₆"],
                                                 units=["(GPa⁻¹·m⁻¹)", "(GPa⁻¹·m⁻¹)", "(GPa⁻¹·m⁻¹)"])
        self.off_axis_a_matrix_entries = off_axis_a_matrix_entries  # Save q_matrix_entries for later use in calculations

        off_axis_D_matrix_entries = self.setup_matrix(self.off_axis_D_tab, "D", row_labels=["M₁", "M₂", "M₆"],
                                                 column_labels=["k₁", "k₂", "k₆"],
                                                 units=["N·m", "N·m", "N·m"])
        self.off_axis_D_matrix_entries = off_axis_D_matrix_entries  # Save q_matrix_entries for later use in calculations

        off_axis_d_matrix_entries = self.setup_matrix(self.off_axis_d_tab, "d", row_labels=["k₁", "k₂", "k₆"],
                                                 column_labels=["M₁", "M₂", "M₆"],
                                                 units=["(N·m⁻¹)", "(N·m⁻¹)", "(N·m⁻¹)"])
        self.off_axis_d_matrix_entries = off_axis_d_matrix_entries  # Save q_matrix_entries for later use in calculations

    def toggle_core(self):
        if self.has_core_var.get():
            self.core_thickness_entry["state"] = "normal"
        else:
            self.core_thickness_entry["state"] = "disabled"
            self.core_thickness_var.set("")  # Clear core thickness if has_core is unchecked

    def set_stress(self):
        if self.N_1_var.get():
            self.N_1 = float(self.N_1_var.get())
        else:
            self.N_1 = 0

        if self.N_2_var.get():
            self.N_2 = float(self.N_2_var.get())
        else:
            self.N_2 = 0

        if self.N_6_var.get():
            self.N_6 = float(self.N_6_var.get())
        else:
            self.N_6 = 0

    def set_moment(self):
        if self.M_1_var.get():
            self.M_1 = float(self.M_1_var.get())
        else:
            self.M_1 = 0

        if self.M_2_var.get():
            self.M_2 = float(self.M_2_var.get())
        else:
            self.M_2 = 0

        if self.M_6_var.get():
            self.M_6 = float(self.M_6_var.get())
        else:
            self.M_6 = 0

    def add_to_layup(self):
        ply_type = self.material_var.get()
        if not ply_type:
            messagebox.showerror("Error", "Please select a Material.")
            return
        thickness_str = self.thickness_var.get()
        orientation_str = self.orientation_var.get()
        has_core = self.has_core_var.get()
        core_thickness_str = self.core_thickness_var.get()

        try:
            thickness = float(thickness_str)
            orientation = float(orientation_str)
        except ValueError:
            messagebox.showerror("Error", "Thickness and Orientation must be numbers.")
            return

        if has_core:
            try:
                core_thickness = float(core_thickness_str)
                if core_thickness < 0:
                    messagebox.showerror("Error", "Core Thickness must be greater than or equal to zero.")
                    return
            except ValueError:
                messagebox.showerror("Error", "Core Thickness must be a number.")
                return
        else:
            core_thickness = None

        unique_id = str(uuid.uuid4())
        ply_number = len(self.data_grid.get_children()) + 1  # Update the count after adding the new layer
        values = (ply_number, ply_type, thickness, orientation, has_core, core_thickness)

        # Use the root item as the parent_item (empty string)
        parent_item = ""
        self.data_grid.insert(parent_item, 0, values=values)
        self.update_ply_numbers()
        self.update_layer_count()  # Update the layer count after adding the new layer

    def move_up(self):
        selected_item = self.data_grid.selection()
        if selected_item:
            current_index = self.data_grid.index(selected_item)
            if current_index > 0:
                self.data_grid.move(selected_item, '', current_index - 1)
                self.update_ply_numbers()

    def move_down(self):
        selected_item = self.data_grid.selection()
        if selected_item:
            current_index = self.data_grid.index(selected_item)
            if current_index < len(self.data_grid.get_children()) - 1:
                self.data_grid.move(selected_item, '', current_index + 1)
                self.update_ply_numbers()

    def copy_symmetric(self):
        # Get all the ply data from the Treeview
        all_ply_data = [(self.data_grid.item(item, 'values'))[1:] for item in self.data_grid.get_children()]

        # Reverse the copied ply data
        symmetric_ply_data = all_ply_data[::-1]

        # Insert new symmetric data into the tree at the correct position
        for index, ply_data in enumerate(symmetric_ply_data):
            ply_number = len(self.data_grid.get_children()) + index + 1
            values = (ply_number,) + tuple(ply_data)

            # Insert at the correct position (index)
            self.data_grid.insert("", index, values=values)

        self.update_ply_numbers()
        self.update_layer_count()

    def delete_selected_row(self):
        selected_item = self.data_grid.selection()
        if selected_item:
            self.data_grid.delete(selected_item)
            self.update_ply_numbers()
            self.update_layer_count()  # Add this line to update the layer count

    def delete_all_entries(self):
        # Ask for confirmation before deleting all entries
        confirmed = messagebox.askyesno("Delete All Entries", "Are you sure you want to delete all entries?")
        if not confirmed:
            return

        # Delete all entries in the Treeview
        for item in self.data_grid.get_children():
            self.data_grid.delete(item)

        # Update ply numbers
        self.update_ply_numbers()
        self.update_layer_count()

    def save_data(self):
        # TODO
        pass

    def calculate(self):
        # Check to make sure there is actually a layup
        self.update_layer_count()

        if self.layer_count <= 0:
            # Display popup message
            messagebox.showerror("Error", "There is nothing to calculate.")
            return  # Exit the function

        # Layup properties
        self.calculate_and_update_off_axis_A_and_a()
        self.calculate_and_update_off_axis_D_and_d()

        # Get the selected ply
        selected_item = self.data_grid.selection()

        if not selected_item:
            # Display warning message
            messagebox.showwarning("Warning", "No layer is selected. Only global layup properties have "
                                              "been calculated!")

        # If a ply is selected, calculate its on-axis properties
        if selected_item:
            # Retrieve data from the selected item
            selected_material = self.data_grid.item(selected_item, 'values')[1]
            selected_orientation = self.data_grid.item(selected_item, 'values')[3]
            selected_layer_number = self.data_grid.item(selected_item, 'values')[0]

            # Update the material properties
            material_properties = mp.get_material_properties(selected_material)
            self.update_tree(self.material_properties_tree, material_properties)

            # Update the strength properties
            strength_properties = mp.get_strength_properties(selected_material)
            self.update_tree(self.strength_tree, strength_properties)

            # Calculate on-axis
            self.calculate_and_update_on_axis_Q_and_S(selected_material)

            # Calculate off-axis
            self.calculate_and_update_off_axis_Q_and_S(selected_material, selected_orientation)

            # z-coordinates of each ply
            self.get_ply_z_coordinates()

            if len(self.z_coordinate) <= 0:
                # Display warning message
                messagebox.showwarning("Warning", "Not enough layers to calculate the strain or stress!")
                return  # Exit the function

            # Calculate off-axis strain
            self.calculate_off_axis_strain(selected_layer_number, self.N_1, self.N_2, self.N_6, self.M_1, self.M_2,
                                           self.M_6)

            # Then calculate on-axis strain for the selected layer
            self.calculate_on_axis_strain(selected_orientation)

            # Then calculate the on-axis stress for the selected layer
            self.calculate_on_axis_stress(selected_material)

    def calculate_and_update_on_axis_Q_and_S(self, material):
        material_properties = mp.get_material_properties(material)

        # Check if the required keys are present in the dictionary
        required_keys = ["E_x", "E_y", "E_s", "ν"]
        for key in required_keys:
            if key not in material_properties:
                messagebox.showerror("Error", f"{key} not found in material properties for {material}")
                return

        # Extract relevant properties
        E_x = material_properties["E_x"]
        E_y = material_properties["E_y"]
        E_s = material_properties["E_s"]
        nu = material_properties["ν"]

        # Calculate coefficients for Q matrix
        nu_y = (E_y / E_x) * nu
        nu_x = nu
        m = (1 - nu * nu_y) ** -1

        self.on_axis_Q_matrix = np.array([[m*E_x, m*nu_y*E_x, 0],
                                     [m*nu_y*E_x, m*E_y, 0],
                                     [0, 0, E_s]])

        self.on_axis_S_matrix = np.linalg.inv(self.on_axis_Q_matrix)

        for index, values in np.ndenumerate(self.on_axis_Q_matrix):
            self.update_entry(self.on_axis_Q_matrix_entries[index[0]][index[1]],
                              self.on_axis_Q_matrix[index[0]][index[1]])

        for index, values in np.ndenumerate(self.on_axis_S_matrix):
            self.update_entry(self.on_axis_S_matrix_entries[index[0]][index[1]],
                              self.on_axis_S_matrix[index[0]][index[1]])

    def calculate_and_update_off_axis_Q_and_S(self, material, orientation):
        # Extract Q_xx, Q_yy, Q_xy, and Q_ss from matrix
        Q_xx = self.on_axis_Q_matrix[0][0]
        Q_xy = self.on_axis_Q_matrix[0][1]
        Q_yx = self.on_axis_Q_matrix[1][0]
        Q_yy = self.on_axis_Q_matrix[1][1]
        Q_ss = self.on_axis_Q_matrix[2][2]

        # Calculate U's
        U_1 = (1/8) * (3 * Q_xx + 3 * Q_yy + 2 * Q_xy + 4 * Q_ss)
        U_2 = (1/2) * (Q_xx - Q_yy)
        U_3 = (1/8) * (Q_xx + Q_yy - 2 * Q_xy - 4 * Q_ss)
        U_4 = (1/8) * (Q_xx + Q_yy + 6 * Q_xy - 4 * Q_ss)
        U_5 = (1/8) * (Q_xx + Q_yy + - 2 * Q_xy + 4 * Q_ss)

        # Calculate the angle in radians
        theta = float(orientation) * (np.pi/180.0)

        # Temp matrix
        temp_matrix = np.array([[U_1, np.cos(2 * theta), np.cos(4 * theta)],
                                [U_1, -np.cos(2 * theta), np.cos(4 * theta)],
                                [U_4, 0, -np.cos(4 * theta)],
                                [U_5, 0, -np.cos(4 * theta)],
                                [0, (1/2) * np.sin(2 * theta), np.sin(4 * theta)],
                                [0, (1/2) * np.sin(2 * theta), -np.sin(4 * theta)]])
        temp_vector = np.array([1, U_2, U_3])
        temp_off_axis_Q = np.dot(temp_matrix, temp_vector)

        self.off_axis_Q_matrix = np.array([[temp_off_axis_Q[0], temp_off_axis_Q[2], temp_off_axis_Q[4]],
                                           [temp_off_axis_Q[2], temp_off_axis_Q[1], temp_off_axis_Q[5]],
                                           [temp_off_axis_Q[4], temp_off_axis_Q[5], temp_off_axis_Q[3]]])

        self.off_axis_S_matrix = np.linalg.inv(self.off_axis_Q_matrix)

        for index, values in np.ndenumerate(self.off_axis_Q_matrix):
            self.update_entry(self.off_axis_Q_matrix_entries[index[0]][index[1]],
                              self.off_axis_Q_matrix[index[0]][index[1]])

        for index, values in np.ndenumerate(self.off_axis_S_matrix):
            self.update_entry(self.off_axis_S_matrix_entries[index[0]][index[1]],
                              self.off_axis_S_matrix[index[0]][index[1]])

    def calculate_and_update_off_axis_A_and_a(self):
        # Retrieve material properties from the entire layup
        children = self.data_grid.get_children()
        last_item = children[-1]
        material_properties = mp.get_material_properties(self.data_grid.item(last_item, 'values')[1])

        # Check if the required keys are present in the dictionary
        required_keys = ["E_x", "E_y", "E_s", "ν"]
        for key in required_keys:
            if key not in material_properties:
                messagebox.showerror("Error",
                                     f"{key} not found in material properties for {self.data_grid.item(last_item, 'values')[1]}")
                return

        # Extract relevant properties
        E_x = material_properties["E_x"]
        E_y = material_properties["E_y"]
        E_s = material_properties["E_s"]
        nu = material_properties["ν"]

        # Calculate coefficients for Q matrix
        nu_y = (E_y / E_x) * nu
        nu_x = nu
        m = (1 - nu_x * nu_y) ** -1

        on_axis_Q_matrix = np.array([[m*E_x, m*nu_y*E_x, 0],
                                     [m*nu_y*E_x, m*E_y, 0],
                                     [0, 0, E_s]])

        Q_xx = on_axis_Q_matrix[0][0]
        Q_xy = on_axis_Q_matrix[0][1]
        Q_yx = on_axis_Q_matrix[1][0]
        Q_yy = on_axis_Q_matrix[1][1]
        Q_ss = on_axis_Q_matrix[2][2]

        # Calculate U's
        U_1 = (1/8) * (3 * Q_xx + 3 * Q_yy + 2 * Q_xy + 4 * Q_ss)
        U_2 = (1/2) * (Q_xx - Q_yy)
        U_3 = (1/8) * (Q_xx + Q_yy - 2 * Q_xy - 4 * Q_ss)
        U_4 = (1/8) * (Q_xx + Q_yy + 6 * Q_xy - 4 * Q_ss)
        U_5 = (1/8) * (Q_xx + Q_yy + - 2 * Q_xy + 4 * Q_ss)

        # Initialize V's
        V_1, V_2, V_3, V_4 = 0, 0, 0, 0

        # Initialize height
        layup_height = 0

        for item in self.data_grid.get_children():
            height = float(self.data_grid.item(item, 'values')[2])
            layup_height += height
            orientation = float(self.data_grid.item(item, 'values')[3])
            theta = orientation * (np.pi / 180.0)
            V_1 += np.cos(2 * theta) * height
            V_2 += np.cos(4 * theta) * height
            V_3 += np.sin(2 * theta) * height
            V_4 += np.sin(4 * theta) * height

        if (layup_height > 0):
            V_1_star = V_1/layup_height
            V_2_star = V_2/layup_height
            V_3_star = V_3/layup_height
            V_4_star = V_4/layup_height
        else:
            V_1_star, V_2_star, V_3_star, V_4_star = 0, 0, 0, 0

        temp_matrix = np.array([[U_1, V_1_star, V_2_star],
                                [U_1, -V_1_star, V_2_star],
                                [U_4, 0, -V_2_star],
                                [U_5, 0, -V_2_star],
                                [0, (1/2) * V_3_star, V_4_star],
                                [0, (1/2) * V_3_star, -V_4_star]])
        temp_vector = np.array([1, U_2, U_3])

        temp_off_axis_A_star = np.dot(temp_matrix, temp_vector)
        temp_off_axis_A = temp_off_axis_A_star * layup_height * (10**(-3))

        self.off_axis_A_matrix = np.array([[temp_off_axis_A[0], temp_off_axis_A[2], temp_off_axis_A[4]],
                                           [temp_off_axis_A[2], temp_off_axis_A[1], temp_off_axis_A[5]],
                                           [temp_off_axis_A[4], temp_off_axis_A[5], temp_off_axis_A[3]]])

        self.off_axis_a_matrix = np.linalg.inv(self.off_axis_A_matrix)

        for index, values in np.ndenumerate(self.off_axis_A_matrix):
            self.update_entry(self.off_axis_A_matrix_entries[index[0]][index[1]],
                              self.off_axis_A_matrix[index[0]][index[1]])

        for index, values in np.ndenumerate(self.off_axis_a_matrix):
            self.update_entry(self.off_axis_a_matrix_entries[index[0]][index[1]],
                              self.off_axis_a_matrix[index[0]][index[1]])

    def calculate_and_update_off_axis_D_and_d(self):
        # Retrieve material properties from the entire layup
        children = self.data_grid.get_children()
        last_item = children[-1]
        material_properties = mp.get_material_properties(self.data_grid.item(last_item, 'values')[1])

        # Check if the required keys are present in the dictionary
        required_keys = ["E_x", "E_y", "E_s", "ν"]
        for key in required_keys:
            if key not in material_properties:
                messagebox.showerror("Error",
                                     f"{key} not found in material properties for {self.data_grid.item(last_item, 'values')[1]}")
                return

        # Extract relevant properties
        E_x = material_properties["E_x"]
        E_y = material_properties["E_y"]
        E_s = material_properties["E_s"]
        nu = material_properties["ν"]

        # Calculate coefficients for Q matrix
        nu_y = (E_y / E_x) * nu
        nu_x = nu
        m = (1 - nu_x * nu_y) ** -1

        on_axis_Q_matrix = np.array([[m*E_x, m*nu_y*E_x, 0],
                                     [m*nu_y*E_x, m*E_y, 0],
                                     [0, 0, E_s]])

        Q_xx = on_axis_Q_matrix[0][0]
        Q_xy = on_axis_Q_matrix[0][1]
        Q_yx = on_axis_Q_matrix[1][0]
        Q_yy = on_axis_Q_matrix[1][1]
        Q_ss = on_axis_Q_matrix[2][2]

        # Calculate U's
        U_1 = (1/8) * (3 * Q_xx + 3 * Q_yy + 2 * Q_xy + 4 * Q_ss)
        U_2 = (1/2) * (Q_xx - Q_yy)
        U_3 = (1/8) * (Q_xx + Q_yy - 2 * Q_xy - 4 * Q_ss)
        U_4 = (1/8) * (Q_xx + Q_yy + 6 * Q_xy - 4 * Q_ss)
        U_5 = (1/8) * (Q_xx + Q_yy + - 2 * Q_xy + 4 * Q_ss)

        # Initialize V's
        V_1, V_2, V_3, V_4 = 0, 0, 0, 0

        # Get core thickness
        if not self.core_thickness_var.get():
            core_thickness = 0
        else:
            core_thickness = float(self.core_thickness_var.get())

        # Initialize height
        half_layup_height = core_thickness

        z_previous = core_thickness
        z_current = z_previous

        for item in self.data_grid.get_children()[int(len(self.data_grid.get_children())/2):]:
            height = float(self.data_grid.item(item, 'values')[2])
            half_layup_height += height
            z_previous = z_current
            z_current += height
            orientation = float(self.data_grid.item(item, 'values')[3])
            theta = orientation * (np.pi / 180.0)
            V_1 += (2/3) * np.cos(2 * theta) * (z_current ** 3 - z_previous ** 3)
            V_2 += (2/3) * np.cos(4 * theta) * (z_current ** 3 - z_previous ** 3)
            V_3 += (2/3) * np.sin(2 * theta) * (z_current ** 3 - z_previous ** 3)
            V_4 += (2/3) * np.sin(4 * theta) * (z_current ** 3 - z_previous ** 3)

        layup_height = half_layup_height * 2
        core_volume_fraction = 2 * core_thickness/layup_height
        h_star = ((layup_height ** 3)/12) * (1 - (core_volume_fraction ** 3))

        temp_matrix = np.array([[U_1, V_1, V_2],
                                [U_1, -V_1, V_2],
                                [U_4, 0, -V_2],
                                [U_5, 0, -V_2],
                                [0, (1/2) * V_3, V_4],
                                [0, (1/2) * V_3, -V_4]])
        temp_vector = np.array([h_star, U_2, U_3])

        temp_off_axis_D = np.dot(temp_matrix, temp_vector)

        self.off_axis_D_matrix = np.array([[temp_off_axis_D[0], temp_off_axis_D[2], temp_off_axis_D[4]],
                                           [temp_off_axis_D[2], temp_off_axis_D[1], temp_off_axis_D[5]],
                                           [temp_off_axis_D[4], temp_off_axis_D[5], temp_off_axis_D[3]]])

        self.off_axis_d_matrix = np.linalg.inv(self.off_axis_D_matrix)

        for index, values in np.ndenumerate(self.off_axis_D_matrix):
            self.update_entry(self.off_axis_D_matrix_entries[index[0]][index[1]],
                              self.off_axis_D_matrix[index[0]][index[1]])

        for index, values in np.ndenumerate(self.off_axis_d_matrix):
            self.update_entry(self.off_axis_d_matrix_entries[index[0]][index[1]],
                              self.off_axis_d_matrix[index[0]][index[1]])

    def calculate_off_axis_strain(self, layer_number, N_1, N_2, N_6, M_1, M_2, M_6):
        # Calculate strain due to in-plane stress
        N_vector = np.array([N_1, N_2, N_6])

        self.in_plane_off_axis_strain_vector = np.dot(self.off_axis_a_matrix, N_vector) * (10**(-9))

        M_vector = np.array([M_1, M_2, M_6])

        curvature_vector = np.dot(self.off_axis_d_matrix, M_vector)

        if self.layer_side_var.get() == "Outer":
            offset = 0
        elif self.layer_side_var.get() == "Inner":
            offset = self.selected_layer_thickness[int(layer_number) - 1]
            print(offset)
        elif self.layer_side_var.get() == "Middle":
            offset = self.selected_layer_thickness[int(layer_number) - 1]/2.0

        current_z_coordinate = self.z_coordinate[int(layer_number) - 1]

        z_coordinate_of_interest = current_z_coordinate

        if current_z_coordinate <= 0:
            z_coordinate_of_interest = current_z_coordinate + offset
        elif current_z_coordinate > 0:
            z_coordinate_of_interest = current_z_coordinate - offset

        self.bending_off_axis_strain_vector = z_coordinate_of_interest * curvature_vector / 1000

        self.off_axis_strain = self.in_plane_off_axis_strain_vector + self.bending_off_axis_strain_vector

        self.update_tree(self.strain_off_tree, {"epsilon_x": "{:.{sf}e}".format(self.off_axis_strain[0], sf=3 - 1),
                "epsilon_y": "{:.{sf}e}".format(self.off_axis_strain[1], sf=3 - 1),
                "epsilon_s": "{:.{sf}e}".format(self.off_axis_strain[2], sf=3 - 1)})

    def calculate_on_axis_strain(self, orientation):
        epsilon_1 = self.off_axis_strain[0]
        epsilon_2 = self.off_axis_strain[1]
        epsilon_6 = self.off_axis_strain[2]

        p = (1 / 2) * (epsilon_1 + epsilon_2)
        q = (1 / 2) * (epsilon_1 - epsilon_2)
        r = (1 / 2) * epsilon_6

        theta = float(orientation) * (np.pi / 180)

        self.epsilon_x = p + q * np.cos(2 * theta) + r * np.sin(2 * theta)
        self.epsilon_y = p - q * np.cos(2 * theta) - r * np.sin(2 * theta)
        self.epsilon_s = - 2 * q * np.sin(2 * theta) + 2 * r * np.cos(2 * theta)

        self.update_tree(self.on_axis_strain_tree, {"epsilon_x": "{:.{sf}e}".format(self.epsilon_x, sf=3 - 1),
                "epsilon_y": "{:.{sf}e}".format(self.epsilon_y, sf=3 - 1),
                "epsilon_s": "{:.{sf}e}".format(self.epsilon_s, sf=3 - 1)})


    def calculate_on_axis_stress(self, material):
        material_properties = mp.get_material_properties(material)

        # Check if the required keys are present in the dictionary
        required_keys = ["E_x", "E_y", "E_s", "ν"]
        for key in required_keys:
            if key not in material_properties:
                messagebox.showerror("Error", f"{key} not found in material properties for {material}")
                return

        # Extract relevant properties
        E_x = material_properties["E_x"]
        E_y = material_properties["E_y"]
        E_s = material_properties["E_s"]
        nu = material_properties["ν"]

        # Calculate coefficients for Q matrix
        nu_y = (E_y / E_x) * nu
        nu_x = nu
        m = (1 - nu * nu_y) ** -1

        self.on_axis_Q_matrix = np.array([[m*E_x, m*nu_y*E_x, 0],
                                     [m*nu_y*E_x, m*E_y, 0],
                                     [0, 0, E_s]]) * (10**(9))

        epsilon_on_axis = np.array([self.epsilon_x, self.epsilon_y, self.epsilon_s])
        sigma_on_axis = np.dot(self.on_axis_Q_matrix, epsilon_on_axis) * (10**(-6))

        self.update_tree(self.on_axis_stress_tree, {"sigma_x": "{:.{sf}e}".format(sigma_on_axis[0], sf=3 - 1),
                "sigma_y": "{:.{sf}e}".format(sigma_on_axis[1], sf=3 - 1),
                "sigma_s": "{:.{sf}e}".format(sigma_on_axis[2], sf=3 - 1)})

    def setup_matrix(self, parent, matrix_name, row_labels=None, column_labels=None, units=None):
        # Matrix Labels
        if row_labels:
            for i, label in enumerate(row_labels):
                ttk.Label(parent, text=label).grid(row=i + 1, column=0, padx=5, pady=5, sticky="e")

        if column_labels:
            for i, label in enumerate(column_labels):
                ttk.Label(parent, text=label).grid(row=0, column=i + 1, padx=5, pady=5, sticky="w")

        # Matrix Entry Widgets
        matrix_entries = []
        for i in range(3):
            row_entries = []
            for j in range(3):
                entry = ttk.Entry(parent, justify="right", width=8)

                # Set the initial value to 0
                entry.insert(0, "0")

                entry.grid(row=i + 1, column=j + 1, padx=5, pady=5)

                # Set the state to readonly after setting the initial value
                entry.config(state="readonly")

                row_entries.append(entry)

            matrix_entries.append(row_entries)

        # Include units in labels
        if units:
            for i in range(3):
                ttk.Label(parent, text=units[i]).grid(row=i + 1, column=4, padx=5, pady=5, sticky="w")

        return matrix_entries

    def update_layer_count(self):
        self.layer_count = len(self.data_grid.get_children())

        # Update the tab label with the layer count
        current_tab = self.layup_notebook.index(self.layup_notebook.select())
        self.layup_notebook.tab(current_tab, text=f"Layer Count: {self.layer_count}")

    def update_ply_numbers(self):
        for index, item in enumerate(self.data_grid.get_children()):
            ply_number = len(self.data_grid.get_children()) - index
            self.data_grid.item(item, values=(ply_number,) + self.data_grid.item(item, 'values')[1:])

    def get_ply_z_coordinates(self):
        # Get core thickness
        if not self.core_thickness_var.get():
            core_thickness = 0
        else:
            core_thickness = float(self.core_thickness_var.get())

        cumulative_thickness = -core_thickness

        num_layers_half = len(self.data_grid.get_children()) // 2
        self.z_coordinate = np.zeros(num_layers_half * 2)
        self.selected_layer_thickness = np.zeros(num_layers_half * 2)

        # Calculate the index for the center layer
        center_index = num_layers_half - 1

        # Iterate over the first half of the layup in reverse
        for index, item in enumerate(reversed(self.data_grid.get_children()[:num_layers_half])):
            thickness = self.data_grid.item(item, 'values')[2]
            self.selected_layer_thickness[index] = float(thickness)
            # Calculate z-coordinate relative to the center of the layup
            cumulative_thickness -= float(thickness)
            self.z_coordinate[index] = cumulative_thickness

        # Reverse the z-coordinate array
        self.z_coordinate[:num_layers_half] = np.flip(self.z_coordinate[:num_layers_half])

        # Concatenate the negative of the array
        self.z_coordinate[num_layers_half:] = -np.flip(self.z_coordinate[:num_layers_half])

    def update_tree(self, tree, data):
        # Clear existing data
        for item in tree.get_children():
            tree.delete(item)

        # Insert new data into the tree
        tree.insert("", "end", values=tuple(data.values()))

    def update_entry(self, entry, value, significant_figures=3):
        # Set the state to normal, update the entry, and set it back to readonly
        entry.config(state="normal")

        if value != 0:
            formatted_value = "{:.{sf}e}".format(value, sf=significant_figures - 1)
        else:
            formatted_value = "0"

        entry.delete(0, tk.END)
        entry.insert(0, formatted_value)
        entry.config(state="readonly")


if __name__ == "__main__":
    root = tk.Tk()
    app = CompositeMaterialsApp(root)
    root.mainloop()
