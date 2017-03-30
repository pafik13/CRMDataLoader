using System;

namespace CRMLite.Entities
{
	/// <summary>
	/// Статус аптеки.
	/// </summary>
	public enum PharmacyState
	{
		psActive, psReserve, psClose
	}

	public class Pharmacy
	{
		public string UUID { get; set; }

		public string State { get; set; }

		public void SetState(PharmacyState newState) { State = newState.ToString("G"); }

		public PharmacyState GetState() { return (PharmacyState)Enum.Parse(typeof(PharmacyState), State, true); }

		public string Brand { get; set; }

		public string NumberName { get; set; }

		public string LegalName { get; set; }

		public string GetName() { return string.Format("{0}, {1}", Brand, Address); }

		public string Net { get; set; }

		public string Address { get; set; }

		public string Subway { get; set; }

		public string Region { get; set; }

		public string Phone { get; set; }

		public string Place { get; set; }

		public string Category { get; set; }

		public int? TurnOver { get; set; }

		public string Comment { get; set; }

		public string Email { get; set; }

		public string CreatedBy { get; set; }

		public DateTimeOffset CreatedAt { get; set; }

		public DateTimeOffset UpdatedAt { get; set; }
	}
}
