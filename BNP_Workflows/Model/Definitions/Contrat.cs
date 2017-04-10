using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using Microsoft.Xrm.Sdk;

namespace BNP_Workflows.Model.Definitions
{
	public class Contrat
	{
		public Programme Programme { get; set; }
		public acquereur acquereur { get; set; }
		public lot lot { get; set; }
		public TMA TMA { get; set; }
		public commissions commissions { get; set; }
	}

	public class Programme
	{
		public string program_name { get; set; }
		public string program_id_gestcom { get; set; }
		public string program_id_wincom { get; set; }
	}

	public class acquereur
	{
		public string firstname { get; set; }
	}

	public class lot
	{
		public string type { get; set; }
	}

	public class TMA
	{
		public string Nom { get; set; }
		public string designation { get; set; }
		public decimal cout { get; set; }
		public string amenager { get; set; }
		public string est_offert { get; set; }
		public bool is_est_offert { get; set; }
		public string Reference { get; set; }
	}

	public class commissions
	{
		public string Nom { get; set; }
		public decimal part_versee_a_la_reservation { get; set; }
		public decimal part_versee_a_l_acte { get; set; }
		public decimal prime_a_l_acte { get; set; }
		public decimal Taux_de_repartition_percent { get; set; }
		public decimal Total_variable { get; set; }
	}
}
