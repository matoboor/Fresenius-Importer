/*
 * Created by SharpDevelop.
 * User: Martin Boor
 * Date: 06.04.2017
 * Time: 10:43
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Globalization;
using System.Text;

namespace Fresenius_Importer
{
	/// <summary>
	/// Description of Model.
	/// </summary>
	public class Model
	{	String mzdroj = "PE";
		String mevidc;
		DateTime mdatdod;
		float mvaha;		
		int mmnoz;
		int mmnoz2;
		String mmj = "COLLI";
		String mnazov;
		String mulica;
		String mmiesto;
		String mpsc;
		String mstat = "SK";
		
		
		//identifikacia zakaznika v L-wise
		public String MZDROJ { 
			get {
				return mzdroj;
			} 
			private set {
				mzdroj = "PE";
			}
		}
		//cislo dodacieho listu
		public String MEVIDC {
			get{
				return mevidc;
			}
			set{
				mevidc = value;
			}
		}
		//datum dodania
		public DateTime MDATDOD {
			get{
				return mdatdod;
			}
			set{
				mdatdod = value;
			}
		}		
		//datum prepravy
		public DateTime MDATPRE {
			get{
				return DateTime.Today;
			}
		}		
		//vaha
		public float MVAHA {
			get{
				if (mvaha<1)
				{
					return 1;
				}
				else
				{
					return mvaha;
				}
			}
			set{
				mvaha = value;
			}
		}
		
		//pocet jednotiek
		public int MMNOZ {
			get{
				return mmnoz;
			}
			set{				
				mmnoz = value;
			}
		}
		
		//pocet paliet
		public int MMNOZ2 {
			get{
				return mmnoz2;
			}
			private set{
				mmnoz2 = value;
			}
		}
		//COLLI
		public String MMJ {
			get{
				return mmj;
			}
			set{
				mmj = "COLLI";
			}
		}
		//nazov prijemcu
		public String MNAZOV {
			get{
				return mnazov.Substring(0, Math.Min(mnazov.Length, 30));
			}
			set{
				mnazov = RemoveDiacritics(value);
			}
		}
		//ulica prijemcu
		public String MULICA {
			get{
				return mulica.Substring(0, Math.Min(mulica.Length, 30));
			}
			set{
				mulica = RemoveDiacritics(value);
			}
		}
		//mesto prijemcu
		public String MMIESTO {
			get{
				return mmiesto.Substring(0, Math.Min(mmiesto.Length, 30));
			}
			set{
				mmiesto = RemoveDiacritics(value);
			}
		}
		//PSC prijemcu
		public String MPSC {
			get{
				return mpsc;
			}
			set{
				mpsc = RemoveSpaces(value);
			}
		}
		//Stat prijemcu - stale SK
		public String MSTAT {
			get { 
				return mstat;
			}
			set {
				mstat = "SK";
			}
			
		}
		//constructor		
		public Model(String evidc, String datdod, String vaha, String mnoz, String nazov, String ulica, String miesto, String psc)
		{
			mevidc = evidc;
			string[] ddd = datdod.Split('.');
			DateTime d = new DateTime(Int32.Parse(ddd[2].Split(' ')[0]),Int32.Parse(ddd[1]),Int32.Parse(ddd[0]));
			//mdatdod = Convert.ToDateTime(datdod.Split(' ')[0], CultureInfo.InvariantCulture);
			//mdatdod = DateTime.ParseExact(dd,"dd.M.yyyy",null);
			mdatdod = d;
			var culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
			culture.NumberFormat.NumberDecimalSeparator = ",";
			
			mvaha = float.Parse(vaha, culture);
			//ak je mnozstvo mensie ako 0.1 -> 1 COLLI
			
			float i = float.Parse(mnoz, culture);
			if (i<0.1) {
				mmnoz = 1;
				mmnoz2 = 0;	
			}
			else
			{
				int pal = (int)Math.Truncate(i);
				float p = i - pal;
				if (p >= 0.1) {
					pal++;
				}
				mmnoz = pal;
				mmnoz2 = pal;
				
			}
			mnazov = RemoveDiacritics(nazov);
			mulica = RemoveDiacritics(ulica);
			mmiesto = RemoveDiacritics(miesto);
			mpsc = RemoveSpaces(psc);
		}

		String RemoveDiacritics(String value)
		{
			var normalizedString = value.Normalize(NormalizationForm.FormD);
			var stringBuilder = new StringBuilder();
    		foreach (var c in normalizedString)
    		{
        		var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
        		if (unicodeCategory != UnicodeCategory.NonSpacingMark)
        		{
           			stringBuilder.Append(c);
        		}
    		}
    		return stringBuilder.ToString().Normalize(NormalizationForm.FormC).ToUpper();
		}

		String RemoveSpaces(String value)
		{
			return value.Replace(" ", String.Empty);
		}
	}
}
