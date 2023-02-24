using ClosedXML.Excel;
using OpsAccountingWF.DataModel;
using System.Data;
using System.Reflection;

namespace OpsAccountingWF
{
	public static class GeneralFunction
	{
		public static DataTable GetDataFromExcel(string path)
		{
			DataTable dt = new DataTable();
			try
			{
				using (XLWorkbook workBook = new XLWorkbook(path))
				{

					IXLWorksheet workSheet = workBook.Worksheet(1);
					bool firstRow = true;
					int skiprows = 1;
					foreach (IXLRow row in workSheet.Rows())
					{
						skiprows = skiprows - 1;
						if (skiprows <= 0)
						{
							//Use the first row to add columns to DataTable.
							if (firstRow)
							{
								int j = 0;
								foreach (IXLCell cell in row.Cells())
								{
									if (!string.IsNullOrEmpty(cell.Value.ToString()))
									{
										dt.Columns.Add(cell.Value.ToString());
									}
									else
									{    //string A = "A" + j;
										 //dt.Columns.Add(A.ToString());
									}
									j++;
								}
								firstRow = false;
							}
							else
							{
								if (!row.IsEmpty())
								{
									int i = 0;
									DataRow toInsert = dt.NewRow();
									foreach (IXLCell cell in row.Cells(1, dt.Columns.Count))
									{
										try
										{
											if (cell.Value.ToString() == "")
											{
												toInsert[i] = "";
											}
											else
											{
												toInsert[i] = cell.Value.ToString();
											}
										}
										catch (Exception ex)
										{
										}
										i++;
									}
									dt.Rows.Add(toInsert);
								}
							}
						}
					}
				}
			}
			catch (Exception ex)
			{

			}
			return dt;
		}

		public static List<EDIDetail> GetListFromExcel(string path)
		{
			//DataTable dt = new DataTable();
			List<EDIDetail> eDIDetails = new List<EDIDetail>();
			Type fields = typeof(EDIDetail);

			PropertyInfo[] props = fields.GetProperties(BindingFlags.Public | BindingFlags.Instance);
			try
			{
				using (XLWorkbook workBook = new XLWorkbook(path))
				{
					IXLWorksheet workSheet = workBook.Worksheet(1);
					int skiprows = 2;
					skiprows = skiprows - 1;

					foreach (IXLRow row in workSheet.Rows())
					{
						if (skiprows <= 0)
						{
							List<IXLCell> cells = row.Cells().ToList();
							//if (cells.Count() != props.Length -1)
							//{
							//	break;
							//}
							//else
							{
								int j = 0;
								EDIDetail eDIDetail = new EDIDetail();
								eDIDetail.LegalEntity = workSheet.Cell(row.RowNumber(),1).Value.ToString();
								eDIDetail.Account_Code = workSheet.Cell(row.RowNumber(), 2).Value.ToString();
								eDIDetail.Supplier_Name = workSheet.Cell(row.RowNumber(), 3).Value.ToString();
								eDIDetail.File_Number = workSheet.Cell(row.RowNumber(), 4).Value.ToString();
								eDIDetail.BL = workSheet.Cell(row.RowNumber(), 5).Value.ToString();
								eDIDetail.Type = workSheet.Cell(row.RowNumber(), 6).Value.ToString();
								eDIDetail.Payref = workSheet.Cell(row.RowNumber(), 7).Value.ToString(); ;
								eDIDetail.Invoice_date = workSheet.Cell(row.RowNumber(), 8).Value.ToString(); ;
								eDIDetail.Date_Create = Convert.ToDateTime(workSheet.Cell(row.RowNumber(), 9).Value.ToString());
								eDIDetail.DESCRIPTION = workSheet.Cell(row.RowNumber(), 10).Value.ToString(); 
								eDIDetail.ARInvCreatedBy = workSheet.Cell(row.RowNumber(), 11).Value.ToString();
								eDIDetail.AMOUNT = Convert.ToDouble(workSheet.Cell(row.RowNumber(), 12).Value.ToString());
								//eDIDetail.AMOUNT = workSheet.Cell(row.RowNumber(), 12).Value.ToString();
								eDIDetail.Currency = workSheet.Cell(row.RowNumber(), 13).Value.ToString();
								eDIDetail.CODE = workSheet.Cell(row.RowNumber(), 14).Value.ToString();
								eDIDetail.Rejection = workSheet.Cell(row.RowNumber(), 15).Value.ToString();
								eDIDetail.SuppRef = workSheet.Cell(row.RowNumber(), 16).Value.ToString();
								eDIDetail.ETA_ETD = workSheet.Cell(row.RowNumber(), 17).Value.ToString(); 
								eDIDetail.ApproveUser = workSheet.Cell(row.RowNumber(), 18).Value.ToString();
								eDIDetail.Application = workSheet.Cell(row.RowNumber(), 19).Value.ToString(); 
								eDIDetail.last_Modify_Date = workSheet.Cell(row.RowNumber(), 20).Value != "" ? Convert.ToDateTime(workSheet.Cell(row.RowNumber(), 20).Value.ToString()) : null;
								eDIDetails.Add(eDIDetail);
							}
							
						}
						skiprows = skiprows - 1;

					}
				}
			}
			catch (Exception ex) { }

			return eDIDetails;
		}
	}
}
