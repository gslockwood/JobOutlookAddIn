using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Utilities
{
	public class Entity
	{
		public string Item { get; private set; }
		public Int32 Attrib { get; private set; }

		public Entity(string item, int attrib)
		{
			Item = item ?? throw new ArgumentNullException(nameof(item));
			Attrib = attrib;
		}

		public Entity(string jobTitle, string level)
		{
			Item = jobTitle ?? throw new ArgumentNullException(nameof(jobTitle));
			Attrib = Convert.ToInt32( level );
		}

		public override string ToString()
		{
			return String.Format( "{0} {1}", this.Item, this.Attrib );
		}
		//
	}
	//
}
