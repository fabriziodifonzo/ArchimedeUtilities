using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excelCsvConverter.model;
using FileHelpers;
using Microsoft.Office.Interop.Excel;

namespace excelCsvConverter
{
	class MainApp
	{
		static void Main(string[] args)
		{
			Microsoft.Office.Interop.Excel.Application wFile = null;
			Microsoft.Office.Interop.Excel.Workbook wBook = null;

			try
			{
				if (_checkArguments(args) != CHECKARGUMENTS_OK)
				{
					Environment.Exit(CHECKARGUMENTS_KO);
				}

				string inputFileName = _extractFileName(args[0]);
				if (_checkExcelFileNameIsOk(inputFileName) != CHECKARGUMENTS_OK)
				{
					Environment.Exit(CHECKARGUMENTS_KO);
				}

				wFile = new Microsoft.Office.Interop.Excel.Application();
				wBook = wFile.Workbooks.Open(inputFileName, ReadOnly: true, Password: "DWD@G");

				Worksheet sheet = (Worksheet)wBook.Sheets[1];
				Microsoft.Office.Interop.Excel.Range xlRange = sheet.UsedRange;
				string newFileName = System.IO.Directory.GetCurrentDirectory() + "\\DataMigration.csv";
				wBook.SaveAs(newFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

				/*
				int rowCount = xlRange.Rows.Count;

				for (int i = 2; i <= rowCount; i++)
				{
					string dipartimento = (string)(xlRange.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2;
					string direzione = (string)(xlRange.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2;
					string ufficio = (string)(xlRange.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value2;
					string areaGiuridica = (string)(xlRange.Cells[i, 4] as Microsoft.Office.Interop.Excel.Range).Value2;
					string cognome = (string)(xlRange.Cells[i, 5] as Microsoft.Office.Interop.Excel.Range).Value2;
					string nome = (string)(xlRange.Cells[i, 6] as Microsoft.Office.Interop.Excel.Range).Value2;
					string codiceFiscale = (string)(xlRange.Cells[i, 7] as Microsoft.Office.Interop.Excel.Range).Value2;
					string indirizzoMail = (string)(xlRange.Cells[i, 8] as Microsoft.Office.Interop.Excel.Range).Value2;
					string tipoMobilita = (string)(xlRange.Cells[i, 9] as Microsoft.Office.Interop.Excel.Range).Value2;
					string dataCessazione = (string)(xlRange.Cells[i, 10] as Microsoft.Office.Interop.Excel.Range).Value2;
					string causaleCessazione = (string)(xlRange.Cells[i, 11] as Microsoft.Office.Interop.Excel.Range).Value2;

					var excelTowData = Employee.Of(dipartimento, direzione, ufficio, areaGiuridica, cognome, nome, codiceFiscale, indirizzoMail, tipoMobilita, dataCessazione, causaleCessazione);

				}
				*/
			}
			finally
			{
				if (wBook != null)
				{
					wBook.Close();
				}

				if (wFile != null)
				{
					wFile.Quit();
				}				 
			}

			Environment.Exit(CHECKARGUMENTS_OK);
		}

		private static int _checkExcelFileNameIsOk(string aExcelFileName)
		{
			Debug.Assert(!string.IsNullOrEmpty(aExcelFileName));

			if (!File.Exists(aExcelFileName))
			{
				Console.Out.WriteLine("Insert a valid excel file name.");
				_printUsage();
				return CHECKARGUMENTS_KO;
			}
			return CHECKARGUMENTS_OK;
		}

		private static int _checkArguments(string[] args)
		{
			if (args.Length == 0)
			{
				System.Console.WriteLine("Please insert excel file name parameter!");
				_printUsage();
				return CHECKARGUMENTS_KO;
			}
			if (args.Length > 1)
			{
				System.Console.WriteLine("Wrong Parameter number!");
				_printUsage();
				return CHECKARGUMENTS_KO;
			}
			string param = args[0];
			if (!param.StartsWith("--excelfilename"))
			{
				System.Console.WriteLine("Wrong Parameter name!");
				_printUsage();
				return CHECKARGUMENTS_KO;
			}
			return CHECKARGUMENTS_OK;
		}

		private static string _extractFileName(string param)
		{
			Debug.Assert(!string.IsNullOrEmpty(param));
			return param.Substring(PARAMETER_PORT.Length + 1);
		}

		private static void _printUsage()
		{
			System.Console.WriteLine("Usage: excelCsvConverter --excelfilename=[EXCELFILENAMEWITHPATH]");
			System.Console.WriteLine("i.e. excelCsvConverter --excelfilename=C:\\\\myexceldir\\\\myexcel.xls");
		}

		private const int CHECKARGUMENTS_OK = 0;
		private const int CHECKARGUMENTS_KO = 1;
		private const string PARAMETER_PORT = "--excelfilename";
	}
}
