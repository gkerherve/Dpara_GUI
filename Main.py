import wx
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
import os


class OtherCalc:
    @staticmethod
    def smooth_and_differentiate(x_values, y_values, smooth_width=2.0, pre_smooth=1, diff_width=1.0, post_smooth=1, algorithm="Gaussian"):
        from scipy.ndimage import gaussian_filter
        from scipy.signal import savgol_filter, wiener
        import numpy as np

        # Smoothing function based on algorithm
        def apply_smooth(data, width, algorithm):
            if algorithm == "Gaussian":
                return gaussian_filter(data, width)
            elif algorithm == "Savitsky-Golay":
                window = int(width * 10) if int(width * 10) % 2 == 1 else int(width * 10) + 1
                return savgol_filter(data, window, 3)
            elif algorithm == "Moving Average":
                window = int(width * 10)
                return np.convolve(data, np.ones(window) / window, mode='same')
            elif algorithm == "Wiener":
                return wiener(data, int(width * 10))
            else:  # "None"
                return data

        # Pre-smoothing passes
        smoothed = y_values.copy()
        for _ in range(int(pre_smooth)):
            smoothed = apply_smooth(smoothed, smooth_width, algorithm)

        # Calculate derivative
        derivative = -1 * np.gradient(smoothed, x_values)

        # Post-smoothing passes
        for _ in range(int(post_smooth)):
            derivative = apply_smooth(derivative, diff_width, algorithm)

        # Normalize derivative to data range
        data_range = np.max(y_values) - np.min(y_values)
        deriv_range = np.max(derivative) - np.min(derivative)
        normalized_deriv = ((derivative - np.min(derivative)) / deriv_range * data_range) + np.min(y_values)

        return normalized_deriv


def set_consistent_fonts(window):
    """Set consistent fonts across the application"""
    font = wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)

    def set_font_recursive(widget):
        widget.SetFont(font)
        for child in widget.GetChildren():
            set_font_recursive(child)

    set_font_recursive(window)


class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="D-Parameter Measurement Tool", size=(1200, 700))

        # Initialize data
        self.x_values = np.array([])
        self.y_values = np.array([])
        self.current_file = None
        self.peak_count = 0
        self.Data = {'Core levels': {}}

        # Create menu bar
        self.create_menu_bar()

        # Create main panel
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour(wx.Colour(240, 240, 240))

        # Create components
        self.create_components()
        self.create_layout()

        # Create mock grid
        self.peak_params_grid = MockGrid()

    def create_menu_bar(self):
        menubar = wx.MenuBar()

        # File menu
        file_menu = wx.Menu()
        open_item = file_menu.Append(wx.ID_OPEN, '&Open\tCtrl+O', 'Open Excel file')
        save_item = file_menu.Append(wx.ID_SAVE, '&Save\tCtrl+S', 'Save to Excel file')
        file_menu.AppendSeparator()
        exit_item = file_menu.Append(wx.ID_EXIT, 'E&xit\tCtrl+Q', 'Exit application')

        # Help menu
        help_menu = wx.Menu()
        help_item = help_menu.Append(wx.ID_HELP, '&Help\tF1', 'Help')

        menubar.Append(file_menu, '&File')
        menubar.Append(help_menu, '&Help')

        self.SetMenuBar(menubar)

        # Bind menu events
        self.Bind(wx.EVT_MENU, self.on_open, open_item)
        self.Bind(wx.EVT_MENU, self.on_save, save_item)
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)
        self.Bind(wx.EVT_MENU, self.on_help, help_item)

    def create_components(self):
        # Create matplotlib figure and canvas for the plot
        self.figure = Figure(figsize=(8, 6))
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvas(self.panel, -1, self.figure)

        # File controls
        self.sheet_label = wx.StaticText(self.panel, label="Select Sheet:")
        self.sheet_combobox = wx.ComboBox(self.panel, style=wx.CB_READONLY)
        self.sheet_combobox.Bind(wx.EVT_COMBOBOX, self.on_sheet_changed)
        self.sheet_combobox.Enable(False)

        # D-Parameter Controls (moved from DParameterWindow)
        # Parameters Box
        self.param_box = wx.StaticBox(self.panel, label="D-Parameter Analysis")

        # Smooth Width
        self.smooth_label = wx.StaticText(self.panel, label="Smooth Width")
        self.smooth_spin = wx.SpinCtrlDouble(self.panel, value="7.0", min=0.1, max=10.0, inc=0.1)

        # Pre-Smooth Passes
        self.pre_label = wx.StaticText(self.panel, label="Pre-Smooth Passes")
        self.pre_spin = wx.SpinCtrl(self.panel, value="2", min=0, max=10)

        # Differentiation Width
        self.diff_label = wx.StaticText(self.panel, label="Differentiation Width (eV)")
        self.diff_spin = wx.SpinCtrlDouble(self.panel, value="1.0", min=0.1, max=10.0, inc=0.1)

        # Post-Smooth Passes
        self.post_label = wx.StaticText(self.panel, label="Post-Smooth Passes")
        self.post_spin = wx.SpinCtrl(self.panel, value="1", min=0, max=10)

        # Smooth Algorithm
        self.algo_label = wx.StaticText(self.panel, label="Smooth Algorithm")
        algorithms = ["Gaussian", "Savitsky-Golay", "Moving Average", "Wiener", "None"]
        self.algo_combo = wx.ComboBox(self.panel, choices=algorithms, style=wx.CB_READONLY)
        self.algo_combo.SetSelection(0)

        # D-Parameter Value Box
        self.d_box = wx.StaticBox(self.panel, label="D-Parameter Value (eV)")
        self.d_value = wx.TextCtrl(self.panel, value="0", style=wx.TE_READONLY)
        self.d_unit = wx.StaticText(self.panel, label="eV")

        # Buttons
        self.clear_btn = wx.Button(self.panel, label="Clear D-para")
        self.calc_btn = wx.Button(self.panel, label="Calculate")
        self.clear_btn.SetMinSize((125, 30))
        self.calc_btn.SetMinSize((125, 30))

        # Initially disable D-parameter controls
        self.enable_d_parameter_controls(False)

        # Bind events
        self.calc_btn.Bind(wx.EVT_BUTTON, self.on_calculate)
        self.clear_btn.Bind(wx.EVT_BUTTON, self.on_clear)

    def enable_d_parameter_controls(self, enable):
        """Enable or disable D-parameter controls"""
        controls = [
            self.smooth_spin, self.pre_spin, self.diff_spin, self.post_spin,
            self.algo_combo, self.clear_btn, self.calc_btn
        ]
        for control in controls:
            control.Enable(enable)

    def create_layout(self):
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Left side - controls (vertical layout)
        left_sizer = wx.BoxSizer(wx.VERTICAL)

        # File controls
        left_sizer.Add(self.sheet_label, 0, wx.ALL, 5)
        left_sizer.Add(self.sheet_combobox, 0, wx.ALL | wx.EXPAND, 5)

        # Add some spacing
        left_sizer.Add(wx.StaticLine(self.panel), 0, wx.ALL | wx.EXPAND, 5)

        # D-Parameter controls
        param_sizer = wx.StaticBoxSizer(self.param_box, wx.VERTICAL)

        # Create grid for parameters
        param_grid_sizer = wx.FlexGridSizer(rows=5, cols=2, hgap=5, vgap=5)
        param_grid_sizer.AddGrowableCol(1)

        param_grid_sizer.Add(self.smooth_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        param_grid_sizer.Add(self.smooth_spin, 0, wx.ALL | wx.EXPAND, 2)

        param_grid_sizer.Add(self.pre_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        param_grid_sizer.Add(self.pre_spin, 0, wx.ALL | wx.EXPAND, 2)

        param_grid_sizer.Add(self.diff_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        param_grid_sizer.Add(self.diff_spin, 0, wx.ALL | wx.EXPAND, 2)

        param_grid_sizer.Add(self.post_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        param_grid_sizer.Add(self.post_spin, 0, wx.ALL | wx.EXPAND, 2)

        param_grid_sizer.Add(self.algo_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 2)
        param_grid_sizer.Add(self.algo_combo, 0, wx.ALL | wx.EXPAND, 2)

        param_sizer.Add(param_grid_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # D-Parameter Value Box
        d_sizer = wx.StaticBoxSizer(self.d_box, wx.HORIZONTAL)
        d_sizer.Add(self.d_value, 1, wx.ALL | wx.EXPAND, 5)
        d_sizer.Add(self.d_unit, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        # Buttons
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_sizer.Add(self.clear_btn, 0, wx.ALL, 2)
        btn_sizer.Add(self.calc_btn, 0, wx.ALL, 2)

        # Add to param sizer
        param_sizer.Add(d_sizer, 0, wx.ALL | wx.EXPAND, 5)
        param_sizer.Add(btn_sizer, 0, wx.ALL | wx.ALIGN_CENTER, 5)

        # Add to left sizer
        left_sizer.Add(param_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # Add stretch to push everything to top
        left_sizer.AddStretchSpacer()

        # Add to main sizer
        main_sizer.Add(left_sizer, 0, wx.ALL | wx.EXPAND, 10)
        main_sizer.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 10)

        self.panel.SetSizer(main_sizer)

    def on_open(self, event):
        """Open Excel file"""
        wildcard = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"

        with wx.FileDialog(self, "Open Excel file", wildcard=wildcard,
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return

            self.current_file = dialog.GetPath()
            self.load_excel_file()

    def load_excel_file(self):
        """Load data from Excel file"""
        try:
            # Read Excel file
            excel_file = pd.ExcelFile(self.current_file)

            # Update sheet selector
            self.sheet_combobox.Clear()
            for sheet_name in excel_file.sheet_names:
                self.sheet_combobox.Append(sheet_name)

            if excel_file.sheet_names:
                self.sheet_combobox.SetSelection(0)
                self.load_sheet_data(excel_file.sheet_names[0])

            self.sheet_combobox.Enable(True)
            self.enable_d_parameter_controls(True)

        except Exception as e:
            wx.MessageBox(f"Error loading Excel file: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def load_sheet_data(self, sheet_name):
        """Load data from specific sheet - columns A and B"""
        try:
            # Read data from columns A and B with headers
            df = pd.read_excel(self.current_file, sheet_name=sheet_name, usecols=[0, 1], header=0)

            # Store column names as axis labels
            self.x_label = df.columns[0]
            self.y_label = df.columns[1]

            # Remove any rows with NaN values
            df = df.dropna()

            if len(df) < 2:
                wx.MessageBox("Sheet must contain at least 2 data points in columns A and B",
                              "Error", wx.OK | wx.ICON_ERROR)
                return

            # Extract x and y values
            self.x_values = df.iloc[:, 0].values
            self.y_values = df.iloc[:, 1].values

            # Initialize data structure for this sheet
            if sheet_name not in self.Data['Core levels']:
                self.Data['Core levels'][sheet_name] = {'Fitting': {'Peaks': {}}}

            # Plot the data
            self.plot_data()

        except Exception as e:
            wx.MessageBox(f"Error loading sheet data: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def on_sheet_changed(self, event):
        """Handle sheet selection change"""
        sheet_name = self.sheet_combobox.GetValue()
        if sheet_name:
            self.load_sheet_data(sheet_name)

    def plot_data(self):
        """Plot the current data"""
        self.ax.clear()

        if len(self.x_values) > 0:
            self.ax.plot(self.x_values, self.y_values, 'b-', label='Data')
            self.ax.set_xlabel(self.x_label if hasattr(self, 'x_label') else 'X Values')
            self.ax.set_ylabel(self.y_label if hasattr(self, 'y_label') else 'Y Values')
            self.ax.set_title('Data Plot')
            self.ax.invert_xaxis()  # Add this line to invert X-axis
            self.ax.legend()
            self.ax.grid(True)

        self.canvas.draw()

    def on_clear(self, event):
        """Clear D-parameter analysis"""
        sheet_name = self.sheet_combobox.GetValue()

        # Clear the data
        if sheet_name in self.Data['Core levels']:
            if 'Fitting' in self.Data['Core levels'][sheet_name]:
                self.Data['Core levels'][sheet_name]['Fitting'] = {}

        # Reset peak count
        self.peak_count = 0

        # Reset D-value display
        self.d_value.SetValue("0")

        # Clear the derivative from plot
        lines_to_remove = []
        for line in self.ax.lines:
            if line.get_label() == 'Derivative':
                lines_to_remove.append(line)

        for line in lines_to_remove:
            line.remove()

        self.canvas.draw_idle()

    def on_calculate(self, event):
        """Calculate D-parameter"""
        if len(self.x_values) == 0:
            wx.MessageBox("No data available for calculation.", "Error", wx.OK | wx.ICON_ERROR)
            return

        x_values = self.x_values
        y_values = self.y_values
        sheet_name = self.sheet_combobox.GetValue()

        # Clear previous derivative plots
        lines_to_remove = []
        for line in self.ax.lines:
            if line.get_label() == 'Derivative':
                lines_to_remove.append(line)

        for line in lines_to_remove:
            line.remove()

        algorithm = self.algo_combo.GetString(self.algo_combo.GetSelection())
        normalized_deriv = OtherCalc.smooth_and_differentiate(
            x_values,
            y_values,
            self.smooth_spin.GetValue(),
            self.pre_spin.GetValue(),
            self.diff_spin.GetValue(),
            self.post_spin.GetValue(),
            algorithm
        )

        min_idx = np.argmin(normalized_deriv)
        max_idx = np.argmax(normalized_deriv)
        x_min = x_values[min_idx]
        x_max = x_values[max_idx]
        center = (x_max + x_min) / 2
        d_param = round(abs(x_max - x_min), 2)

        # Update D-value display
        self.d_value.SetValue(f"{d_param:.2f}")

        # Update Data dictionary
        if 'Fitting' not in self.Data['Core levels'][sheet_name]:
            self.Data['Core levels'][sheet_name]['Fitting'] = {}
        if 'Peaks' not in self.Data['Core levels'][sheet_name]['Fitting']:
            self.Data['Core levels'][sheet_name]['Fitting']['Peaks'] = {}

        peak_data = {
            'Position': center,
            'FWHM': d_param,
            'Sigma': self.pre_spin.GetValue(),
            'Gamma': self.post_spin.GetValue(),
            'Skew': self.smooth_spin.GetValue(),
            'L/G': self.diff_spin.GetValue(),
            'Fitting Model': 'D-parameter',
            'Derivative': normalized_deriv.tolist()
        }

        self.Data['Core levels'][sheet_name]['Fitting']['Peaks']['D-parameter'] = peak_data

        # Update plot
        self.ax.plot(x_values, normalized_deriv, '-', color='red', label='Derivative', linewidth=1)
        # self.ax.invert_xaxis()  # Add this line to maintain inverted axis
        self.ax.legend()
        self.canvas.draw_idle()

    def on_save(self, event):
        """Save to Excel file"""
        if not self.current_file:
            wildcard = "Excel files (*.xlsx)|*.xlsx"

            with wx.FileDialog(self, "Save Excel file", wildcard=wildcard,
                               style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dialog:

                if dialog.ShowModal() == wx.ID_CANCEL:
                    return

                self.current_file = dialog.GetPath()

        try:
            current_sheet = self.sheet_combobox.GetValue()

            # Read existing Excel file to preserve other sheets
            existing_sheets = {}
            if os.path.exists(self.current_file):
                existing_file = pd.ExcelFile(self.current_file)
                for sheet in existing_file.sheet_names:
                    if sheet != current_sheet:  # Keep other sheets unchanged
                        existing_sheets[sheet] = pd.read_excel(self.current_file, sheet_name=sheet)

            # Prepare data for current sheet
            df_data = pd.DataFrame({
                self.x_label if hasattr(self, 'x_label') else 'X': self.x_values,
                self.y_label if hasattr(self, 'y_label') else 'Y': self.y_values
            })

            # Add derivative data if it exists
            if (current_sheet in self.Data['Core levels'] and
                    'Fitting' in self.Data['Core levels'][current_sheet] and
                    'Peaks' in self.Data['Core levels'][current_sheet]['Fitting'] and
                    'D-parameter' in self.Data['Core levels'][current_sheet]['Fitting']['Peaks']):

                d_data = self.Data['Core levels'][current_sheet]['Fitting']['Peaks']['D-parameter']
                if 'Derivative' in d_data:
                    df_data['Derivative'] = d_data['Derivative']

            # Write all sheets to Excel
            with pd.ExcelWriter(self.current_file, engine='openpyxl') as writer:
                # Write current sheet with updated data
                df_data.to_excel(writer, sheet_name=current_sheet, index=False)

                # Write back all other existing sheets
                for sheet_name, sheet_data in existing_sheets.items():
                    sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)

            wx.MessageBox("File saved successfully!", "Save Complete", wx.OK | wx.ICON_INFORMATION)

        except Exception as e:
            wx.MessageBox(f"Error saving file: {str(e)}", "Save Error", wx.OK | wx.ICON_ERROR)

    def on_exit(self, event):
        """Exit application"""
        self.Close()

    def on_help(self, event):
        """Show help dialog"""
        help_text = """D-Parameter Measurement Tool

1. Open Excel file (File -> Open)
2. Select sheet with data in columns A and B
3. Adjust D-Parameter analysis parameters
4. Click Calculate to perform analysis
5. Save results (File -> Save)

Parameters:
- Smooth Width: Controls smoothing strength
- Pre/Post-Smooth Passes: Number of smoothing iterations
- Differentiation Width: Width for derivative calculation
- Algorithm: Smoothing method
"""
        wx.MessageBox(help_text, "Help", wx.OK | wx.ICON_INFORMATION)


class MockGrid:
    """Simple mock to replace the grid functionality"""

    def GetNumberRows(self):
        return 0

    def DeleteRows(self, pos, num):
        pass

    def AppendRows(self, num):
        pass

    def SetCellValue(self, row, col, value):
        pass

    def GetCellValue(self, row, col):
        return ""

    def GetNumberCols(self):
        return 14

    def SetCellBackgroundColour(self, row, col, color):
        pass


class DParameterApp(wx.App):
    def OnInit(self):
        frame = MainFrame()
        frame.Show()
        return True


if __name__ == '__main__':
    app = DParameterApp()
    app.MainLoop()