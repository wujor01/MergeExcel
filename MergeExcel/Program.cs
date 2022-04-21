using Aspose.Cells;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace MergeExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.OutputEncoding = Encoding.UTF8;
                Console.Write("Đường dẫn file cách nhau bằng dấu ',': ");
                string input = Console.ReadLine().Replace('"', ' ');
                Console.Write("Số cột: ");
                string col = Console.ReadLine().Replace('"', ' ');

                Console.WriteLine($"Đang xử lý merge file...");

                var dt = ListExcelToDatatable(input.Split(','), int.Parse(col));
                var file = ExportDatatableToExcel(dt);

                string filename = $"{Guid.NewGuid().ToString()}.xlsx";

                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), filename);

                File.WriteAllBytes(path, file);
                Console.WriteLine($"Xuất file thành công file được lưu ở ngoài màn hình desktop với tên {filename}");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }

        }

        public static DataTable ListExcelToDatatable(string[] listPath, int countCols)
        {
            try
            {
				DataTable dataTable = new DataTable();

                foreach (var path in listPath)
                {
					Stream stream = File.OpenRead(path.Trim());

                    string filename = path.Split('\\').ToList().LastOrDefault();

                    // Instantiate a Workbook object
                    //Opening the Excel file through the file stream
                    Workbook workbook = new Workbook(stream);


					foreach (var worksheet in workbook.Worksheets)
					{
                        try
                        {
                            int countRows = worksheet.Cells.Rows.Count;

                            if (countRows > 2)
                            {
                                DataTable dtNew = worksheet.Cells.ExportDataTable(0, 0, countRows, countCols, true);
                                DataColumn Col = dtNew.Columns.Add("SHEET", System.Type.GetType("System.String"));
                                Col.SetOrdinal(0);

                                foreach (DataRow dr in dtNew.Rows)
                                {
                                    dr["SHEET"] = worksheet.Name;
                                }

                                dataTable.Merge(dtNew);
                            }
						}
                        catch (Exception ex)
                        {
							Console.WriteLine($"Lỗi file {filename} sheet {worksheet.Name}");
							Console.WriteLine(ex.ToString());
						}

						
					}

					// Close the file stream to free all resources
					stream.Close();
				}

				return dataTable;

			}
            catch (Exception)
            {
                throw;
            }
        }

        public static byte[] ExportDatatableToExcel(DataTable dtData)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage pck = new ExcelPackage())
            {
                //Create the worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet1");

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                ws.Cells["A1"].LoadFromDataTable(dtData, true);

                //Format the header for column 1-3
                using (OfficeOpenXml.ExcelRange rng = ws.Cells["A1:BZ1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;   //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                }

                for (int i = 0; i < dtData.Columns.Count; i++)
                {
                    if (dtData.Columns[i].DataType == typeof(DateTime))
                    {
                        using (OfficeOpenXml.ExcelRange col = ws.Cells[2, i + 1, 2 + dtData.Rows.Count, i + 1])
                        {
                            //col.Style.Numberformat.Format = "MM/dd/yyyy HH:mm";
                            col.Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
                            //col.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        }
                    }
                    if (dtData.Columns[i].DataType == typeof(TimeSpan))
                    {
                        using (OfficeOpenXml.ExcelRange col = ws.Cells[2, i + 1, 2 + dtData.Rows.Count, i + 1])
                        {
                            col.Style.Numberformat.Format = "d.hh:mm";
                            col.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }
                    }
                }

                return pck.GetAsByteArray();
            } // end using

        }
    }
}
