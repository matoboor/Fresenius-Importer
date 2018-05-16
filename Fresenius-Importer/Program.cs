/*
 * Created by SharpDevelop.
 * User: Martin Boor
 * Date: 04.04.2017
 * Time: 10:03
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Excel;

namespace Fresenius_Importer
{
	class Program
	{
		
		public static void Main(string[] args)
		{
			
			// PARAMETERS
			String sourceDBfilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"source"));
			String sourceExcelFilePath = @"C:\Fresenius-import";
			String targetLwisPath = @"C:\Fresenius_edi";
			
			// COUNTERS
			int lines = 0;
			int files = 0;
			
			List<Model> records = new List<Model>();
			DataSet result;
			DataTable table;
			
			//scan source path for excel import file
			string[] importFiles = new string[] { };
			try
			{	
				importFiles = Directory.GetFiles(sourceExcelFilePath,"*.XLSX");
				Console.WriteLine("Počet nájdených súborov: {0}", importFiles.Length);
				if (importFiles.Length > 0) {
					Console.WriteLine(importFiles[0]);
				
					Console.WriteLine("Čítam zdrojový súbor: {0}", Path.GetFileName(importFiles[0]));
					FileStream fs = File.Open(importFiles[0], FileMode.Open, FileAccess.Read);
					IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
					result = reader.AsDataSet();
					table = result.Tables[0];
					foreach (DataRow dr in table.Rows) {
						if (lines > 0 && lines != table.Rows.Count) {
							Model m = new Model(dr[0].ToString(), dr[7].ToString(), dr[5].ToString(), dr[6].ToString(), dr[1].ToString(), dr[2].ToString(), dr[4].ToString(), dr[3].ToString());
							records.Add(m);
						}
						lines++;
					}	
					fs.Close();
				}
				else
				{
					Console.ForegroundColor = ConsoleColor.Red;
					Console.WriteLine("Nenašiel som žiadny súbor!!!");
					Console.ResetColor();
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("Chyba v čítaní zdrojového priečinka!!! : {0} ", e.Message);
			}
			
			Console.WriteLine("Počet záznamov v importnom súbore: {0}", lines-1);
			Console.WriteLine("Počet nahraných záznamov: {0}", records.Count);
			
			if (records.Count > 0) {
			
				Console.WriteLine("Vytváram súbory DBF");
			
				Writer w = new Writer();
				if (w.WriteFile(records)) {
					w.CleanDirectories(sourceExcelFilePath);
					Console.WriteLine("Hotovo");
				
				}
			
				Console.ForegroundColor = ConsoleColor.Green;
				Console.WriteLine("Zuzanka usmej sa :) a stlač lubovoľné tlačidlo");
			
				Console.ReadKey(true);
			}
			else
			{
				Console.ForegroundColor = ConsoleColor.Red;
				Console.WriteLine("Niečo sa posralo!!! Nepreklínaj ma - volaj mi");
			
				Console.ReadKey(true);
			}
		}
	}
}