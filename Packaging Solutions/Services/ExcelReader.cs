using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using PackagingSolutions.Model;

namespace PackagingSolutions.Services
{
    public class ExcelReader
    {
        public List<PackageElement> ReadExcel(string filePath)
        {
            List<PackageElement> elements = new List<PackageElement>();

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Error: The specified file does not exist.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return elements;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        MessageBox.Show("Error: The Excel file contains no worksheets.", "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return elements;
                    }

                    var worksheet = package.Workbook.Worksheets[0]; // First sheet
                    int row = 2; // Row 1 is header

                    while (worksheet.Cells[row, 1].Value != null)
                    {
                        string elementType = worksheet.Cells[row, 1].Text.Trim();
                        string name = worksheet.Cells[row, 2].Text.Trim();

                        if (!string.IsNullOrEmpty(elementType) && !string.IsNullOrEmpty(name))
                        {
                            elements.Add(new PackageElement
                            {
                                ElementType = elementType,
                                Name = name
                            });
                        }

                        row++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading Excel file: " + ex.Message, "Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return elements;
        }
    }
}
