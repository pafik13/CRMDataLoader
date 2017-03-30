using System;

namespace CRMLite.Entities
{
	public enum Sex
	{
		Male, Female
	}

	public class Employee
	{
		public string UUID { get; set; }

		public string Pharmacy { get; set; }

		public string Name { get; set; }

		public string Sex { get; set; }

		public void SetSex(Sex newSex) { Sex = newSex.ToString("G"); }

		public Sex GetSex() { return (Sex)Enum.Parse(typeof(Sex), Sex, true); }

		public string Position { get; set; }

		public bool IsCustomer { get; set; }

		public DateTimeOffset? BirthDate { get; set; }

		public string Phone { get; set; }

		public string Email { get; set; }

		public string Loyalty { get; set; }

		public bool CanParticipate { get; set; }

		public string Comment { get; set; }

        public string CreatedBy { get; set; }

        public DateTimeOffset CreatedAt { get; set; }

		public DateTimeOffset UpdatedAt { get; set; }

		public bool IsSynced { get; set; }
	}
}

