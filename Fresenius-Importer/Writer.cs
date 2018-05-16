/*
 * Created by SharpDevelop.
 * User: Martin Boor
 * Date: 11.04.2017
 * Time: 15:18
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Fresenius_Importer
{
	/// <summary>
	/// Description of Writer.
	/// </summary>
	public class Writer
	{
		String sourceDBLfilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"source\\L.dbf"));
		String sourceDBPfilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"source\\P.dbf"));
		String tmpFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"tmp"));
		String targetLwisPath = @"G:\PRENOSY\LWIS\VYSTUP";
		String counterFilePath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"counter\\counter.txt"));
		int fileCounter = 0;
		
		public Writer()
		{
		}
		
		public String GetFileName()
		{
			try
			{
				StreamReader sr = new StreamReader(counterFilePath);
				string content = sr.ReadLine();
				int oldCounter = Convert.ToInt32(content);
				oldCounter++;
				int newCounter = oldCounter;
				sr.Close();
				StreamWriter sw = new StreamWriter(counterFilePath);
				sw.Write(newCounter);
				sw.Close();
				return newCounter+".dbf";
			}
			catch(Exception e)
			{
				Console.WriteLine("COUNTER ERROR> " + e.Message);
				return "Error";
			}
		}
		
		public String WriteFile(Model model)
		{
			if (File.Exists(sourceDBLfilePath) && File.Exists(sourceDBPfilePath)) {
				string fileName = GetFileName();
				string tmpL = tmpFilePath + "\\L_4" + fileName;
				string tmpP = tmpFilePath + "\\P_4" + fileName;
				
				File.Copy(sourceDBLfilePath,tmpL , true);
				File.Copy(sourceDBPfilePath,tmpP , true);
				
				string connString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Path.GetDirectoryName(tmpL) + ";Extended Properties=DBASE III;";
				try
				{
					OleDbConnection conn = new OleDbConnection(connString);
					conn.Open();
					
					OleDbCommand cmd = new OleDbCommand();
					
					cmd.CommandType = CommandType.Text;
					cmd.CommandText = "INSERT INTO " + Path.GetFileName(tmpL) + " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
									
					cmd.Parameters.AddWithValue("MZMENA", "");
					cmd.Parameters.AddWithValue("MDOKLAD", "..........");
					cmd.Parameters.AddWithValue("MZDROJ", model.MZDROJ);
					cmd.Parameters.AddWithValue("MEVIDC", model.MEVIDC);
					cmd.Parameters.AddWithValue("MDATDOD", model.MDATDOD);
					cmd.Parameters.AddWithValue("MVAHA", model.MVAHA);
					cmd.Parameters.AddWithValue("MMNOZ", model.MMNOZ);
					cmd.Parameters.AddWithValue("MMJ", model.MMJ);
					cmd.Parameters.AddWithValue("MNAZOV", model.MNAZOV);
					cmd.Parameters.AddWithValue("MULICA", model.MULICA);
					cmd.Parameters.AddWithValue("MMIESTO", model.MMIESTO);
					cmd.Parameters.AddWithValue("MPSC", model.MPSC);
					cmd.Parameters.AddWithValue("MSTAT", model.MSTAT);
					cmd.Parameters.AddWithValue("MDATPRE", model.MDATPRE);
					cmd.Parameters.AddWithValue("MCASPRE", "");
					cmd.Parameters.AddWithValue("MODVDOV", "0");
					cmd.Parameters.AddWithValue("MOBJEM", 0);
					cmd.Parameters.AddWithValue("MCISPRAC", "");
					cmd.Parameters.AddWithValue("MVYSKL", "");
					cmd.Parameters.AddWithValue("MPREPRAV", 0);
					cmd.Parameters.AddWithValue("MTELEFON", "");
					cmd.Parameters.AddWithValue("MSPESNINA", "");
					cmd.Parameters.AddWithValue("MPREPRAV1", 0);
					cmd.Parameters.AddWithValue("MDATNAKL", model.MDATPRE);
					cmd.Parameters.AddWithValue("MTYPMOJ", "");
					cmd.Parameters.AddWithValue("MODOS", "");
					cmd.Parameters.AddWithValue("MSTRED", "");
					cmd.Parameters.AddWithValue("MSKLAD", "");
					cmd.Parameters.AddWithValue("MZONA", "");
					cmd.Parameters.AddWithValue("MSPRAC", "");
					cmd.Parameters.AddWithValue("MDATCASTL", "");
					cmd.Parameters.AddWithValue("MVRATPAL", 0);
					cmd.Parameters.AddWithValue("MSKUTPAL", 0);
					cmd.Parameters.AddWithValue("MICO", " ");
					cmd.Parameters.AddWithValue("MEDITPAL", "");
					cmd.Parameters.AddWithValue("MSTAVPAL", "");
					cmd.Parameters.AddWithValue("MODOS2", "");
					cmd.Parameters.AddWithValue("MODOSU", "");
					cmd.Parameters.AddWithValue("MODOSM", "");
					cmd.Parameters.AddWithValue("MODOSP", "");
					cmd.Parameters.AddWithValue("MIMPORT", "");					
					cmd.Parameters.AddWithValue("MMNOZ2", model.MMNOZ2);
					cmd.Parameters.AddWithValue("MMJ2", "");
					
					cmd.Connection = conn;
					cmd.ExecuteNonQuery();
					Console.WriteLine(Path.GetFileName(tmpL) + "> OK");
					fileCounter++;
					
					OleDbCommand cmd1 = new OleDbCommand();
					
					cmd1.CommandType = CommandType.Text;
					cmd1.CommandText = "INSERT INTO " + Path.GetFileName(tmpP) + " values (?,?,?,?,?)";
					
					cmd1.Parameters.AddWithValue("PZMENA", "");
					cmd1.Parameters.AddWithValue("PZDROJ", model.MZDROJ);
					cmd1.Parameters.AddWithValue("PEVIDC", model.MEVIDC);
					cmd1.Parameters.AddWithValue("PDRUH", "");
					cmd1.Parameters.AddWithValue("PPOZNAM", "");
					
					cmd1.Connection = conn;
					cmd1.ExecuteNonQuery();
					Console.WriteLine(Path.GetFileName(tmpP) + "> OK");
					fileCounter++;
					
					conn.Close();
				}
				catch (Exception e)
				{
					Console.WriteLine(e.Message);
					Console.WriteLine(e.StackTrace);

					return "False";
				}
				return "True";
			}
			else
			{
				return "False";
			}
		}
		
		public bool WriteFile(List<Model> list)
		{
			int r = 0;
			foreach (Model m in list)
			{
				if( WriteFile(m).Equals("True"))
				{
					r++;
				}				   
			}
			CopyFiles();
			return r == list.Count;
			
		}
		
		public void CopyFiles()
		{
			Console.WriteLine("Kopírujem súbory na sieťový disk");
			string[] filesToChangeExtension = Directory.GetFiles(tmpFilePath, "*.dbf");
			foreach (string file in filesToChangeExtension)
			{
				File.Move(file, Path.ChangeExtension(file, ".DAT"));
				Console.Write(".");
			}
			
			string[] filesToCopy = Directory.GetFiles(tmpFilePath, "*.DAT");
			foreach (string file in filesToCopy)
			{
				File.Copy(file, Path.Combine(targetLwisPath, Path.GetFileName(file)));
				Console.Write(".");
			}
			Console.WriteLine();
			
		}
		
		public void CleanDirectories(string inputExcelPath)
		{
			Console.WriteLine("Upratujem po sebe");
			
			DirectoryInfo di = new DirectoryInfo(tmpFilePath);
			foreach (FileInfo file in di.GetFiles())
			{
    			file.Delete();
				Console.Write(".");
			}
			
			DirectoryInfo di1 = new DirectoryInfo(inputExcelPath);
			foreach (FileInfo file in di1.GetFiles())
			{
    			file.Delete();
				Console.Write(".");
			}
			Console.WriteLine();
		}
		
	}
}
