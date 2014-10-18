//  Excel__Module__Generator.cs
//
//  Author:
//       Victor L. Senior (VLS) <betselection(&)gmail.com>
//
//  Web: 
//       http://betselection.cc/betsoftware/
//
//  Sources:
//       http://github.com/betselection/
//
//  Copyright (c) 2014 Victor L. Senior
//
//  This program is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with this program.  If not, see <http://www.gnu.org/licenses/>.

/// <summary>
/// Excel Module Generator.
/// </summary>
namespace Excel__Module__Generator
{
    // Directives
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows.Forms;

    /// <summary>
    /// Excel Module Generator class.
    /// </summary>
    public class Excel__Module__Generator : Form
    {
        /// <summary>
        /// The load excel file button.
        /// </summary>
        private Button loadExcelFileButton;

        /// <summary>
        /// The sheet format combo box.
        /// </summary>
        private ComboBox sheetFormatComboBox;

        /// <summary>
        /// The set sheet format label.
        /// </summary>
        private Label setSheetFormatLabel;

        /// <summary>
        /// The set excel sheet label.
        /// </summary>
        private Label setExcelSheetLabel;

        /// <summary>
        /// The set module type label.
        /// </summary>
        private Label setModuleTypeLabel;

        /// <summary>
        /// The module type list box.
        /// </summary>
        private ListBox moduleTypeListBox;

        /// <summary>
        /// The generate module label.
        /// </summary>
        private Label generateModuleLabel;

        /// <summary>
        /// The compile module button.
        /// </summary>
        private Button compileModuleButton;

        /// <summary>
        /// The excel file.
        /// </summary>
        private string excelFile = string.Empty;

        /// <summary>
        /// The type of the module.
        /// </summary>
        private string moduleType = string.Empty;

        /// <summary>
        /// The marshal object.
        /// </summary>
        private object marshal = null;

        /// <summary>
        /// The main open file dialog.
        /// </summary>
        private OpenFileDialog mainOpenFileDialog = new OpenFileDialog();

        /// <summary>
        /// Initializes a new instance of the <see cref="Excel__Module__Generator.Excel__Module__Generator"/> class.
        /// </summary>
        public Excel__Module__Generator()
        {
            this.loadExcelFileButton = new Button();
            this.setExcelSheetLabel = new Label();
            this.setModuleTypeLabel = new Label();
            this.moduleTypeListBox = new ListBox();
            this.generateModuleLabel = new Label();
            this.compileModuleButton = new Button();
            this.setSheetFormatLabel = new Label();
            this.sheetFormatComboBox = new ComboBox();
            this.SuspendLayout();

            // loadExcelFileButton
            this.loadExcelFileButton.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.loadExcelFileButton.Location = new Point(12, 30);
            this.loadExcelFileButton.Name = "loadExcelFileButton";
            this.loadExcelFileButton.Size = new Size(177, 41);
            this.loadExcelFileButton.TabIndex = 0;
            this.loadExcelFileButton.Text = "Load Excel File";
            this.loadExcelFileButton.UseVisualStyleBackColor = true;
            this.loadExcelFileButton.Click += new System.EventHandler(this.LoadExcelFileButtonClick);

            // setExcelSheetLabel
            this.setExcelSheetLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.setExcelSheetLabel.Location = new Point(12, 9);
            this.setExcelSheetLabel.Name = "setExcelSheetLabel";
            this.setExcelSheetLabel.Size = new Size(177, 23);
            this.setExcelSheetLabel.TabIndex = 1;
            this.setExcelSheetLabel.Text = "1) Set Excel Sheet:";

            // setModuleTypeLabel
            this.setModuleTypeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.setModuleTypeLabel.Location = new Point(12, 136);
            this.setModuleTypeLabel.Name = "setModuleTypeLabel";
            this.setModuleTypeLabel.Size = new Size(177, 23);
            this.setModuleTypeLabel.TabIndex = 1;
            this.setModuleTypeLabel.Text = "3) Set module type:";

            // moduleTypeListBox
            this.moduleTypeListBox.FormattingEnabled = true;
            this.moduleTypeListBox.Items.AddRange(new object[]
                {
                    "Input",
                    "Bet Selection",
                    "Money Management",
                    "Display",
                    "Output"
                });
            this.moduleTypeListBox.Location = new Point(12, 159);
            this.moduleTypeListBox.Name = "moduleTypeListBox";
            this.moduleTypeListBox.Size = new Size(177, 69);
            this.moduleTypeListBox.TabIndex = 2;
            this.moduleTypeListBox.SelectedIndexChanged += new System.EventHandler(this.ModuleTypeListBoxSelectedIndexChanged);

            // generateModuleLabel
            this.generateModuleLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.generateModuleLabel.Location = new Point(12, 241);
            this.generateModuleLabel.Name = "generateModuleLabel";
            this.generateModuleLabel.Size = new Size(177, 23);
            this.generateModuleLabel.TabIndex = 1;
            this.generateModuleLabel.Text = "4) Generate Module";

            // compileModuleButton
            this.compileModuleButton.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.compileModuleButton.Location = new Point(12, 265);
            this.compileModuleButton.Name = "compileModuleButton";
            this.compileModuleButton.Size = new Size(177, 41);
            this.compileModuleButton.TabIndex = 3;
            this.compileModuleButton.Text = "Compile Module";
            this.compileModuleButton.UseVisualStyleBackColor = true;
            this.compileModuleButton.Click += new System.EventHandler(this.CompileModuleButtonClick);

            // setSheetFormatLabel
            this.setSheetFormatLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, (byte)0);
            this.setSheetFormatLabel.Location = new Point(12, 79);
            this.setSheetFormatLabel.Name = "setSheetFormatLabel";
            this.setSheetFormatLabel.Size = new Size(177, 23);
            this.setSheetFormatLabel.TabIndex = 1;
            this.setSheetFormatLabel.Text = "2) Set sheet format:";

            // sheetFormatComboBox
            this.sheetFormatComboBox.FormattingEnabled = true;
            this.sheetFormatComboBox.Items.AddRange(new object[]
                {
                    "ExcelBot"
                });
            this.sheetFormatComboBox.Location = new Point(13, 103);
            this.sheetFormatComboBox.Name = "sheetFormatComboBox";
            this.sheetFormatComboBox.Size = new Size(176, 21);
            this.sheetFormatComboBox.TabIndex = 1;
            this.sheetFormatComboBox.Text = "ExcelBot";
            this.sheetFormatComboBox.SelectedIndexChanged += new System.EventHandler(this.SheetFormatComboBoxSelectedIndexChanged);

            // MainForm
            this.AutoScaleDimensions = new SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(201, 315);
            this.Controls.Add(this.sheetFormatComboBox);
            this.Controls.Add(this.moduleTypeListBox);
            this.Controls.Add(this.compileModuleButton);
            this.Controls.Add(this.loadExcelFileButton);
            this.Controls.Add(this.generateModuleLabel);
            this.Controls.Add(this.setSheetFormatLabel);
            this.Controls.Add(this.setModuleTypeLabel);
            this.Controls.Add(this.setExcelSheetLabel);
            this.Font = new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, (byte)0);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.Text = "Excel Module Generator";
            this.ResumeLayout(false);

            // mainOpenFileDialog
            this.mainOpenFileDialog.InitialDirectory = Application.StartupPath;
            this.mainOpenFileDialog.Filter = "Excel files (*.xl*)|*.xl*|All files (*.*)|*.*";
            this.mainOpenFileDialog.RestoreDirectory = true;
        }

        /// <summary>
        /// Inits the instance.
        /// </summary>
        /// <param name="passedMarshal">Passed marshal.</param>
        public void Init(object passedMarshal)
        {
            // Set marshal
            this.marshal = passedMarshal;

            // Set icon
            this.Icon = (Icon)this.marshal.GetType().GetProperty("Icon").GetValue(this.marshal, null);

            // Show form
            this.Show();
        }

        /// <summary>
        /// Loads the excel file button click.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void LoadExcelFileButtonClick(object sender, EventArgs e)
        {
            // Open file dialog
            if (this.mainOpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Set file
                this.excelFile = this.mainOpenFileDialog.FileName;
            }
        }

        /// <summary>
        /// Sheets the format combo box selected index changed.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void SheetFormatComboBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            // TODO Add more formats (kattila, etc.)
        }

        /// <summary>
        /// Modules the type list box selected index changed.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void ModuleTypeListBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            // Set module type
            this.moduleType = this.moduleTypeListBox.SelectedItem.ToString();
        }

        /// <summary>
        /// Compiles the module button click.
        /// </summary>
        /// <param name="sender">Sender object.</param>
        /// <param name="e">Event arguments.</param>
        private void CompileModuleButtonClick(object sender, EventArgs e)
        {
            // Check there's an excel file
            if (this.excelFile.Length == 0)
            {
                // Advice user
                MessageBox.Show("Please set an excel file.", "Excel file", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Halt flow
                return;
            }

            // Check there's a valid module type
            if (this.moduleTypeListBox.SelectedIndex == -1)
            {
                // Advice user
                MessageBox.Show("Please set module type.", "Module type", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Halt flow
                return;
            }

            // Compiler path
            string compilerPath = string.Empty;

            // Module namespace
            string moduleNamespace = this.DisplayNameToNameSpace(Path.GetFileNameWithoutExtension(this.excelFile));

            // Fetch paths from marshal
            Dictionary<string, string> marshalPaths = (Dictionary<string, string>)this.marshal.GetType().GetProperty("Paths").GetValue(this.marshal, null);

            // (F)rame(w)ork (lib)rary flag
            bool fwlib = true;

            // Determine compiler path
            foreach (string exeFile in System.IO.Directory.GetFiles(Directory.GetParent(System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory()).ToString(), "*cs*.exe", SearchOption.AllDirectories))
            {
                // Switch possible compiler paths
                switch (Path.GetFileName(exeFile))
                {
                    case "csc.exe": // .net (unified)
                            // Set as current file
                        compilerPath = exeFile;

                            // Disable framework library
                        fwlib = false;

                            // Exit loop
                        break;
                    case "mcs.exe": // mono (unified)
                            // Set as current file
                        compilerPath = exeFile;

                            // Exit loop
                        break;

                    case "dmcs.exe": // mono (4.0 mscorlib)
                            // Set as current file
                        compilerPath = exeFile;

                            // Exit loop
                        break;

                    case "smcs.exe": // mono (2.1 mscorlib)
                            // Set as current file
                        compilerPath = exeFile;

                            // Exit loop
                        break;

                    case "gmcs.exe": // mono (2.0 mscorlib)
                            // Set as current file
                        compilerPath = exeFile;

                            // Exit loop
                        break;
                }
            }

            // Compiler arguments 
            string compilerArguments = "/noconfig /nostdlib /nologo /warn:4 /optimize+ /t:library /out:<module_name>.dll /r:mscorlib.dll /r:System.dll /r:System.Windows.Forms.dll /r:ExcelApi.dll /r:OfficeApi.dll /r:VBIDEApi.dll /r:NetOffice.dll /r:System.Drawing.dll" + (fwlib ? " \"/lib:" + System.Runtime.InteropServices.RuntimeEnvironment.GetRuntimeDirectory() + "\" " : string.Empty) + " <module_name>.cs";

            // Variable for module file string
            string moduleFileString = string.Empty;

            // Reader for resource stream
            using (StreamReader reader = new StreamReader(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Connector.cs")))
            {
                // Set module file string
                moduleFileString = reader.ReadToEnd();
            }

            // Replace <module_name> with actual name
            moduleFileString = moduleFileString.Replace("<module_name>", moduleNamespace);

            // Replace <module_extension> with actual excel file extension
            moduleFileString = moduleFileString.Replace("<module_extension>", Path.GetExtension(this.excelFile));

            // Connector.cs to StringReader
            using (StringReader reader = new StringReader(moduleFileString))
            {
                // Open file with actual name for writing
                    using (StreamWriter writer = new StreamWriter(Path.Combine(marshalPaths["framework"], moduleNamespace + ".cs")))
                {
                    // Variable for current line
                    string currentLine;

                    // Iterate reader lines
                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        // Save line using current platform's newline format
                        writer.Write(currentLine + Environment.NewLine);
                    }
                }
            }

            // Replace in arguments list
            compilerArguments = compilerArguments.Replace("<module_name>", moduleNamespace);

            // Process start info for the compiler
            ProcessStartInfo compilerProcessStartInfo = new ProcessStartInfo();

            // Set working directory (output file location)
            compilerProcessStartInfo.WorkingDirectory = marshalPaths["framework"];

            // Set compiler file name
            compilerProcessStartInfo.FileName = compilerPath;

            // Set compiler arguments
            compilerProcessStartInfo.Arguments = compilerArguments;

            // Use shell
            compilerProcessStartInfo.UseShellExecute = true;

            // No window
            compilerProcessStartInfo.CreateNoWindow = true;

            // Hidden style
            compilerProcessStartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // Start compilation
            Process compilerProcess = Process.Start(compilerProcessStartInfo);

            // Wait until it finishes
            compilerProcess.WaitForExit();

            // Move generated dll module to proper directory
            File.Move(Path.Combine(marshalPaths["framework"], moduleNamespace + ".dll"), Path.Combine(Path.Combine(Path.Combine(marshalPaths["framework"], this.moduleTypeListBox.SelectedItem.ToString().Replace(" ", string.Empty)), (string)this.marshal.GetType().GetProperty("Game").GetValue(this.marshal, null)), moduleNamespace + ".dll"));    

            // Copy renamed excel file to proper directory
            File.Copy(this.excelFile, Path.Combine(Path.Combine(Path.Combine(marshalPaths["framework"], this.moduleTypeListBox.SelectedItem.ToString().Replace(" ", string.Empty)), (string)this.marshal.GetType().GetProperty("Game").GetValue(this.marshal, null)), moduleNamespace + Path.GetExtension(this.excelFile)), true);

            // Remove source code file
            File.Delete(Path.Combine(marshalPaths["framework"], moduleNamespace + ".cs"));

            // Reload modules in framework
            this.marshal.GetType().GetMethod("ReloadModules").Invoke(this.marshal, null);

            // Advice user about successful compilation
            MessageBox.Show("Successful Module Compilation", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Changes passed display name to namespace.
        /// </summary>
        /// <returns>Resulting name space.</returns>
        /// <param name="displayName">Display name.</param>
        private string DisplayNameToNameSpace(string displayName)
        {
            // Check strings are there
            if (displayName.Length > 0)
            {
                // Match with regular expression
                MatchCollection matches = Regex.Matches(displayName, @"[^a-zA-Z0-9_]");

                // Walk reversed
                for (int i = matches.Count - 1; i >= 0; i--)
                {
                    // Handle space
                    if (matches[i].Value == " ")
                    {
                        // Remove original
                        displayName = displayName.Remove(matches[i].Index, 1);

                        // Insert replacement
                        displayName = displayName.Insert(matches[i].Index, "__");

                        // Next iteration
                        continue;
                    }

                    // Set encoding
                    UTF32Encoding encoding = new UTF32Encoding(); 

                    // Get current bytes
                    byte[] bytes = encoding.GetBytes(matches[i].Value.ToCharArray()); 

                    // Remove original
                    displayName = displayName.Remove(matches[i].Index, 1);

                    // Insert replacement
                    displayName = displayName.Insert(matches[i].Index, "_" + BitConverter.ToInt32(bytes, 0).ToString() + "_");
                }

                // Return processed display name
                return displayName;
            }

            // Return empty string by default
            return string.Empty;
        }
    }
}